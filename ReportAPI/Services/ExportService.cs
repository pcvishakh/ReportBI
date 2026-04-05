using System.Text.Json;
using ClosedXML.Excel;
using ReportAPI.Models;

namespace ReportAPI.Services
{
    public class ExportService
    {
        private const string BlankLabel = "(Blank)";
        private readonly IExportConfigParser _configParser;
        private readonly IDataProcessor _dataProcessor;

        public ExportService(IExportConfigParser configParser, IDataProcessor dataProcessor)
        {
            _configParser = configParser;
            _dataProcessor = dataProcessor;
        }

        public byte[] GenerateExcel(IEnumerable<IDictionary<string, object>> data, string? gridStateJson, IEnumerable<ColumnDefinition>? columnDefinitions = null)
        {
            if (data == null || !data.Any()) return Array.Empty<byte>();
            var dataList = data.ToList();

            using var workbook = new XLWorkbook();
            var (styledColumns, layouts) = _configParser.ParseGridState(gridStateJson);

            if (layouts.Any())
            {
                for (int i = 0; i < layouts.Count; i++)
                {
                    var layout = layouts[i];
                    string sheetName = string.IsNullOrWhiteSpace(layout.Name) ? $"Export_{i + 1}" : layout.Name;
                    ProcessLayout(workbook, dataList, layout, sheetName, i + 1, styledColumns, columnDefinitions);
                }
            }
            else
            {
                ProcessLayout(workbook, dataList, new LayoutConfig { Name = "Export" }, "Export", 1, styledColumns, columnDefinitions);
            }

            using var ms = new MemoryStream();
            workbook.SaveAs(ms);
            return ms.ToArray();
        }

        private void ProcessLayout(XLWorkbook workbook, List<IDictionary<string, object>> dataList, LayoutConfig config, string rawSheetName, int sheetIndex, List<StyledColumnDesc> styledColumns, IEnumerable<ColumnDefinition>? columnDefinitions)
        {
            var processedData = _dataProcessor.ProcessData(dataList, config, columnDefinitions);

            if (config.PivotColumns.Any() || config.PivotGroupedColumns.Any())
            {
                GeneratePivotTable(workbook, processedData, config, rawSheetName, sheetIndex);
            }
            else
            {
                GenerateGridSheet(workbook, processedData, config, rawSheetName, sheetIndex, styledColumns, columnDefinitions);
            }
        }

        private string SanitizeSheetName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) name = "Sheet";
            foreach (char c in new[] { '\\', '/', '*', '?', ':', '[', ']' })
                name = name.Replace(c, '_');
            return name.Length > 31 ? name.Substring(0, 31) : name;
        }

        private string EnsureUniqueSheetName(XLWorkbook workbook, string name, int index)
        {
            string currentName = name;
            int suffix = 0;
            while (workbook.Worksheets.Any(x => x.Name.Equals(currentName, StringComparison.OrdinalIgnoreCase)))
            {
                suffix++;
                currentName = SanitizeSheetName($"{name}_{index}_{suffix}");
            }
            return currentName;
        }

        private void GeneratePivotTable(XLWorkbook workbook, List<IDictionary<string, object>> filteredData, LayoutConfig config, string rawSheetName, int sheetIndex)
        {
            string dataSheetName = SanitizeSheetName(rawSheetName + "_Raw");
            string ptSheetName = SanitizeSheetName(rawSheetName);

            dataSheetName = EnsureUniqueSheetName(workbook, dataSheetName, sheetIndex);
            ptSheetName = EnsureUniqueSheetName(workbook, ptSheetName, sheetIndex);
            if (dataSheetName == ptSheetName) dataSheetName += "_Raw";

            var dataSheet = workbook.Worksheets.Add(dataSheetName);
            var firstRow = filteredData.Any() ? filteredData.First() : new Dictionary<string, object>();
            var allKeys = firstRow.Keys.ToList();

            var pivotCols = GetPivotColumns(config, allKeys, firstRow);
            var rowLabelMap = pivotCols.RowCols.ToDictionary(c => c, c => "__RL__" + c);
            var colLabelMap = pivotCols.ColCols.ToDictionary(c => c, c => "__CL__" + c);
            
            var allKeysToInclude = pivotCols.RawVisibleCols
                .Union(rowLabelMap.Values)
                .Union(colLabelMap.Values)
                .Distinct()
                .ToList();

            var rowFrequencies = CalculateHierarchicalFrequencies(filteredData, pivotCols.RowCols);
            var colFrequencies = CalculateHierarchicalFrequencies(filteredData, pivotCols.ColCols);

            WritePivotData(dataSheet, filteredData, allKeysToInclude, pivotCols.RawVisibleCols, pivotCols.DimensionCols, pivotCols.RowCols, pivotCols.ColCols, rowLabelMap, colLabelMap, rowFrequencies, colFrequencies);

            var ptSheet = workbook.Worksheets.Add(ptSheetName);
            if (filteredData.Any() && allKeysToInclude.Any())
            {
                SetupPivotTable(ptSheet, dataSheet, filteredData.Count, allKeysToInclude, config, rowLabelMap, colLabelMap);
            }

            dataSheet.Hide();
            ptSheet.Columns().AdjustToContents();
        }

        private (List<string> RowCols, List<string> ColCols, List<string> AggCols, List<string> DimensionCols, List<string> RawVisibleCols) GetPivotColumns(LayoutConfig config, List<string> allKeys, IDictionary<string, object> firstRow)
        {
            var tableCols = config.TableColumns;
            if (!tableCols.Any() && allKeys.Any())
            {
                tableCols = allKeys.Where(k => !_dataProcessor.IsJson(firstRow[k])).ToList();
            }

            var gridVisibleCols = tableCols.Where(c => !config.ColumnVisibility.TryGetValue(c, out var isVis) || isVis).ToList();
            var rowCols = config.PivotGroupedColumns.Intersect(allKeys).ToList();
            var colCols = config.PivotColumns.Intersect(allKeys).ToList();
            var aggCols = config.PivotAggregationColumns.Select(a => a.ColumnId).Intersect(allKeys).ToList();
            var dimensionCols = rowCols.Union(colCols).ToList();
            var rawVisibleCols = gridVisibleCols.Union(rowCols).Union(colCols).Union(aggCols).Intersect(allKeys).Distinct().ToList();

            return (rowCols, colCols, aggCols, dimensionCols, rawVisibleCols);
        }

        private Dictionary<string, int> CalculateHierarchicalFrequencies(List<IDictionary<string, object>> data, List<string> groupCols)
        {
            var frequencies = new Dictionary<string, int>();
            if (!groupCols.Any()) return frequencies;

            foreach (var row in data)
            {
                var currentPrefix = "";
                for (int i = 0; i < groupCols.Count; i++)
                {
                    var colName = groupCols[i];
                    var val = row.TryGetValue(colName, out var v) && v != null && !string.IsNullOrWhiteSpace(v.ToString()) ? v.ToString() : BlankLabel;
                    currentPrefix = i == 0 ? val! : $"{currentPrefix}|||{val}";
                    
                    if (frequencies.ContainsKey(currentPrefix)) frequencies[currentPrefix]++;
                    else frequencies[currentPrefix] = 1;
                }
            }
            return frequencies;
        }

        private void WritePivotData(IXLWorksheet sheet, List<IDictionary<string, object>> data, List<string> allKeysToInclude, List<string> rawVisibleCols, List<string> dimensionCols, List<string> rowCols, List<string> colCols, Dictionary<string, string> rowMap, Dictionary<string, string> colMap, Dictionary<string, int> rowFreqs, Dictionary<string, int> colFreqs)
        {
            for (int i = 0; i < allKeysToInclude.Count; i++)
            {
                sheet.Cell(1, i + 1).Value = allKeysToInclude[i];
                sheet.Cell(1, i + 1).Style.Font.Bold = true;
            }

            for (int r = 0; r < data.Count; r++)
            {
                var row = data[r];
                for (int c = 0; c < rawVisibleCols.Count; c++)
                {
                    var colName = rawVisibleCols[c];
                    if (row.TryGetValue(colName, out var v) && v != null && !string.IsNullOrWhiteSpace(v.ToString()))
                    {
                        sheet.Cell(r + 2, c + 1).Value = XLCellValue.FromObject(v);
                    }
                    else if (dimensionCols.Contains(colName))
                    {
                        sheet.Cell(r + 2, c + 1).Value = BlankLabel;
                    }
                }

                int labelColOffset = rawVisibleCols.Count;
                var dimensionLabelCols = rowMap.Values.Union(colMap.Values).ToList();
                for (int i = 0; i < dimensionLabelCols.Count; i++)
                {
                    var labelHeader = dimensionLabelCols[i];
                    bool isRowLabel = labelHeader.StartsWith("__RL__");
                    var colName = labelHeader.Substring(6);
                    
                    var hierarchy = isRowLabel ? rowCols : colCols;
                    var freqs = isRowLabel ? rowFreqs : colFreqs;
                    
                    int level = hierarchy.IndexOf(colName);
                    string labelValue;

                    if (level >= 0)
                    {
                        var prefixParts = new List<string>();
                        for (int l = 0; l <= level; l++)
                        {
                            var hCol = hierarchy[l];
                            var val = row.TryGetValue(hCol, out var v) && v != null && !string.IsNullOrWhiteSpace(v.ToString()) ? v.ToString() : BlankLabel;
                            prefixParts.Add(val!);
                        }
                        var fullPrefix = string.Join("|||", prefixParts);
                        var displayVal = prefixParts.Last();
                        labelValue = freqs.TryGetValue(fullPrefix, out int count) ? $"{displayVal} ({count})" : displayVal;
                    }
                    else
                    {
                        var val = row.TryGetValue(colName, out var v) && v != null && !string.IsNullOrWhiteSpace(v.ToString()) ? v.ToString() : BlankLabel;
                        labelValue = val!;
                    }
                    
                    sheet.Cell(r + 2, labelColOffset + i + 1).Value = labelValue;
                }
            }
        }

        private void SetupPivotTable(IXLWorksheet ptSheet, IXLWorksheet dataSheet, int rowCount, List<string> allKeys, LayoutConfig config, Dictionary<string, string> rowMap, Dictionary<string, string> colMap)
        {
            var sourceRange = dataSheet.Range(1, 1, rowCount + 1, allKeys.Count);
            var pt = ptSheet.PivotTables.Add("PivotTable_" + ptSheet.Name.Replace(" ", "_"), ptSheet.Cell(1, 1), sourceRange);

            foreach (var grp in config.PivotGroupedColumns)
            {
                if (rowMap.TryGetValue(grp, out var rl)) pt.RowLabels.Add(rl);
            }
            foreach (var col in config.PivotColumns)
            {
                if (colMap.TryGetValue(col, out var cl)) pt.ColumnLabels.Add(cl);
            }
            foreach (var agg in config.PivotAggregationColumns)
            {
                var field = pt.Values.Add(agg.ColumnId);
                field.CustomName = $"{agg.AggFunc}({agg.ColumnId})";
                field.SummaryFormula = agg.AggFunc.ToLower() switch
                {
                    "sum" => XLPivotSummary.Sum,
                    "count" => XLPivotSummary.Count,
                    "avg" => XLPivotSummary.Average,
                    "min" => XLPivotSummary.Minimum,
                    "max" => XLPivotSummary.Maximum,
                    _ => XLPivotSummary.Count
                };
            }

            pt.Theme = XLPivotTableTheme.PivotStyleMedium9;
        }

        private void GenerateGridSheet(XLWorkbook workbook, List<IDictionary<string, object>> filteredData, LayoutConfig config, string rawSheetName, int sheetIndex, List<StyledColumnDesc> styledColumns, IEnumerable<ColumnDefinition>? columnDefinitions)
        {
            string rsName = EnsureUniqueSheetName(workbook, SanitizeSheetName(rawSheetName), sheetIndex);
            var worksheet = workbook.Worksheets.Add(rsName);

            var visibleColumns = GetVisibleColumns(config, filteredData);
            WriteGridHeaders(worksheet, visibleColumns, config.TableAggregationColumns);

            int currentRow = 2;
            if (config.RowGroupedColumns.Any())
            {
                WriteGroup(worksheet, filteredData, config.RowGroupedColumns, 0, visibleColumns, config.TableAggregationColumns, config.ColumnSorts, ref currentRow, 1, columnDefinitions);
                worksheet.Outline.SummaryVLocation = XLOutlineSummaryVLocation.Top;
            }
            else
            {
                WriteGridData(worksheet, filteredData, visibleColumns, ref currentRow, columnDefinitions);
            }

            worksheet.Columns().AdjustToContents();
            ApplyConditionalFormatting(worksheet, visibleColumns, styledColumns, currentRow);
        }

        private List<string> GetVisibleColumns(LayoutConfig config, List<IDictionary<string, object>> data)
        {
            var tableCols = config.TableColumns;
            if (!tableCols.Any() && data.Any())
            {
                var firstRow = data.First();
                tableCols = firstRow.Keys.Where(k => !_dataProcessor.IsJson(firstRow[k])).ToList();
            }
            tableCols.Remove("ag-Grid-AutoColumn");

            var visibleColumns = new List<string>();
            if (config.RowGroupedColumns.Any()) visibleColumns.Add("Group");
            visibleColumns.AddRange(tableCols.Where(c => !config.ColumnVisibility.TryGetValue(c, out var isVis) || isVis));

            if (!visibleColumns.Any()) visibleColumns.Add("Data");
            return visibleColumns;
        }

        private void WriteGridHeaders(IXLWorksheet sheet, List<string> columns, List<AggregationDesc> aggs)
        {
            for (int i = 0; i < columns.Count; i++)
            {
                string headerName = columns[i];
                var aggDesc = aggs.FirstOrDefault(a => a.ColumnId == headerName);
                if (aggDesc != null && !string.IsNullOrWhiteSpace(aggDesc.AggFunc))
                    headerName = $"{aggDesc.AggFunc}({headerName})";

                var cell = sheet.Cell(1, i + 1);
                cell.Value = headerName;
                cell.Style.Font.Bold = true;
            }
        }

        private void WriteGridData(IXLWorksheet sheet, List<IDictionary<string, object>> data, List<string> columns, ref int currentRow, IEnumerable<ColumnDefinition>? columnDefinitions)
        {
            foreach (var row in data)
            {
                for (int c = 0; c < columns.Count; c++)
                {
                    var colName = columns[c];
                    if (row.TryGetValue(colName, out var val) && val != null)
                    {
                        var cell = sheet.Cell(currentRow, c + 1);
                        cell.Value = XLCellValue.FromObject(val);
                        ApplyFormat(cell, colName, columnDefinitions);
                    }
                }
                currentRow++;
            }
        }

        private void WriteGroup(IXLWorksheet worksheet, IEnumerable<IDictionary<string, object>> data, List<string> groupCols, int groupLevel, List<string> visibleCols, List<AggregationDesc> aggs, List<SortDesc> sorts, ref int currentRow, int outlineLevel, IEnumerable<ColumnDefinition>? columnDefinitions)
        {
            if (groupLevel >= groupCols.Count)
            {
                WriteRawGroupData(worksheet, data, visibleCols, ref currentRow, outlineLevel, columnDefinitions);
                return;
            }

            var col = groupCols[groupLevel];
            var grouped = data.GroupBy(r => r.TryGetValue(col, out var val) && val != null ? val.ToString() ?? BlankLabel : BlankLabel).ToList();

            foreach (var grp in grouped)
            {
                WriteGroupHeader(worksheet, grp, groupLevel, visibleCols, aggs, ref currentRow, outlineLevel, columnDefinitions);
                WriteGroup(worksheet, grp.ToList(), groupCols, groupLevel + 1, visibleCols, aggs, sorts, ref currentRow, outlineLevel + 1, columnDefinitions);
            }
        }

        private void WriteRawGroupData(IXLWorksheet worksheet, IEnumerable<IDictionary<string, object>> data, List<string> visibleCols, ref int currentRow, int outlineLevel, IEnumerable<ColumnDefinition>? columnDefinitions)
        {
            foreach (var row in data)
            {
                for (int c = 1; c < visibleCols.Count; c++)
                {
                    var colName = visibleCols[c];
                    if (row.TryGetValue(colName, out var val) && val != null)
                    {
                        var cell = worksheet.Cell(currentRow, c + 1);
                        cell.Value = XLCellValue.FromObject(val);
                        ApplyFormat(cell, colName, columnDefinitions);
                    }
                }
                if (outlineLevel > 1) worksheet.Row(currentRow).OutlineLevel = outlineLevel - 1;
                currentRow++;
            }
        }

        private void WriteGroupHeader(IXLWorksheet worksheet, IGrouping<string, IDictionary<string, object>> grp, int groupLevel, List<string> visibleCols, List<AggregationDesc> aggs, ref int currentRow, int outlineLevel, IEnumerable<ColumnDefinition>? columnDefinitions)
        {
            string indent = new string(' ', groupLevel * 4);
            worksheet.Cell(currentRow, 1).Value = $"{indent}{grp.Key ?? ""} ({grp.Count()})";

            foreach (var agg in aggs)
            {
                int colIndex = visibleCols.IndexOf(agg.ColumnId);
                if (colIndex >= 0)
                {
                    var colDef = columnDefinitions?.FirstOrDefault(c => c.Field == agg.ColumnId);
                    object? aggResult = _dataProcessor.ComputeAggregate(grp, agg, colDef);
                    if (aggResult != null)
                    {
                        var cell = worksheet.Cell(currentRow, colIndex + 1);
                        cell.Value = XLCellValue.FromObject(aggResult);
                        ApplyFormat(cell, agg.ColumnId, columnDefinitions);
                    }
                }
            }

            worksheet.Row(currentRow).Style.Font.Bold = true;
            if (outlineLevel > 1) worksheet.Row(currentRow).OutlineLevel = outlineLevel - 1;
            currentRow++;
        }

        private void ApplyFormat(IXLCell cell, string colName, IEnumerable<ColumnDefinition>? columnDefinitions)
        {
            var colDef = columnDefinitions?.FirstOrDefault(cd => cd.Field == colName);
            if (colDef?.DataType == "date")
                cell.Style.NumberFormat.Format = "mm/dd/yyyy";
        }

        private void ApplyConditionalFormatting(IXLWorksheet worksheet, List<string> visibleColumns, List<StyledColumnDesc> styledColumns, int currentRow)
        {
            if (currentRow <= 2) return;
            foreach (var styledCol in styledColumns)
            {
                int colIndex = visibleColumns.IndexOf(styledCol.ColumnId);
                if (colIndex >= 0)
                {
                    var range = worksheet.Range(2, colIndex + 1, currentRow - 1, colIndex + 1);
                    foreach (var cr in styledCol.GradientRanges)
                        ApplyGradientRange(range, cr);
                }
            }
        }

        private void ApplyGradientRange(IXLRange range, CellRangeDesc cr)
        {
            if (cr.Min.HasValue && cr.Max.HasValue)
                range.AddConditionalFormat().WhenBetween(cr.Min.Value, cr.Max.Value).Fill.SetBackgroundColor(XLColor.FromHtml(cr.Color));
            else if (cr.Min.HasValue)
                range.AddConditionalFormat().WhenEqualOrGreaterThan(cr.Min.Value).Fill.SetBackgroundColor(XLColor.FromHtml(cr.Color));
            else if (cr.Max.HasValue)
                range.AddConditionalFormat().WhenEqualOrLessThan(cr.Max.Value).Fill.SetBackgroundColor(XLColor.FromHtml(cr.Color));
        }
    }
}
