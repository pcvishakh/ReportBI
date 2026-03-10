using System.Text.Json;
using ClosedXML.Excel;

namespace ReportAPI.Services
{
    public class ExportService
    {
        public byte[] GenerateExcel(IEnumerable<IDictionary<string, object>> data, string? gridStateJson)
        {
            if (data == null) return Array.Empty<byte>();
            var dataList = data.ToList();

            bool layoutProcessed = false;
            int sheetIndex = 0;
            List<StyledColumnDesc> styledColumns = new();

            using var workbook = new XLWorkbook();

            if (!string.IsNullOrWhiteSpace(gridStateJson))
            {
                try
                {
                    var doc = JsonDocument.Parse(gridStateJson);
                    styledColumns = ParseStyledColumns(doc.RootElement);

                    if (doc.RootElement.TryGetProperty("Layout", out var layoutEl) &&
                        layoutEl.TryGetProperty("Layouts", out var layoutsEl) &&
                        layoutsEl.GetArrayLength() > 0)
                    {
                        foreach (var layout in layoutsEl.EnumerateArray())
                        {
                            sheetIndex++;
                            string sheetName = layout.TryGetProperty("Name", out var nameProp) ? nameProp.GetString() ?? $"Export_{sheetIndex}" : $"Export_{sheetIndex}";
                            ProcessLayout(workbook, dataList, layout, sheetName, sheetIndex, styledColumns);
                            layoutProcessed = true;
                        }
                    }
                }
                catch { }
            }

            if (!layoutProcessed)
            {
                ProcessLayout(workbook, dataList, default(JsonElement), "Export", 1, styledColumns);
            }

            using var ms = new MemoryStream();
            workbook.SaveAs(ms);
            return ms.ToArray();
        }

        private List<StyledColumnDesc> ParseStyledColumns(JsonElement root)
        {
            var styledColumns = new List<StyledColumnDesc>();
            if (root.TryGetProperty("StyledColumn", out var scEl) &&
                scEl.TryGetProperty("StyledColumns", out var scsEl))
            {
                foreach (var scItem in scsEl.EnumerateArray())
                {
                    var styledCol = new StyledColumnDesc
                    {
                        ColumnId = scItem.GetProperty("ColumnId").GetString() ?? ""
                    };

                    if (scItem.TryGetProperty("GradientStyle", out var gsEl) &&
                        gsEl.TryGetProperty("CellRanges", out var gcrsEl))
                    {
                        foreach (var crItem in gcrsEl.EnumerateArray())
                        {
                            styledCol.GradientRanges.Add(ParseCellRange(crItem));
                        }
                    }

                    if (scItem.TryGetProperty("PercentBarStyle", out var psEl) &&
                        psEl.TryGetProperty("CellRanges", out var pcrsEl))
                    {
                        foreach (var crItem in pcrsEl.EnumerateArray())
                        {
                            styledCol.PercentBarRanges.Add(ParseCellRange(crItem));
                        }
                    }
                    styledColumns.Add(styledCol);
                }
            }
            return styledColumns;
        }

        private CellRangeDesc ParseCellRange(JsonElement crItem)
        {
            return new CellRangeDesc
            {
                Min = crItem.TryGetProperty("Min", out var minV) ? minV.GetDouble() : (double?)null,
                Max = crItem.TryGetProperty("Max", out var maxV) ? maxV.GetDouble() : (double?)null,
                Color = crItem.TryGetProperty("Color", out var colV) ? colV.GetString() ?? "" : ""
            };
        }

        private string SanitizeSheetName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) name = "Sheet";
            foreach (char c in new[] { '\\', '/', '*', '?', ':', '[', ']' })
                name = name.Replace(c, '_');
            return name.Length > 31 ? name.Substring(0, 31) : name;
        }

        private void ProcessLayout(XLWorkbook workbook, List<IDictionary<string, object>> dataList, JsonElement layoutJson, string rawSheetName, int sheetIndex, List<StyledColumnDesc> styledColumns)
        {
            var config = ParseLayoutConfig(layoutJson);
            var filteredData = ApplyFilters(dataList, config.ColumnFilters);
            var sortedData = ApplySorting(filteredData, config.ColumnSorts);

            if (config.PivotColumns.Any() || config.PivotGroupedColumns.Any())
            {
                GeneratePivotTable(workbook, sortedData, config, rawSheetName, sheetIndex);
            }
            else
            {
                GenerateGridSheet(workbook, sortedData, config, rawSheetName, sheetIndex, styledColumns);
            }
        }

        private LayoutConfig ParseLayoutConfig(JsonElement layoutJson)
        {
            var config = new LayoutConfig();
            if (layoutJson.ValueKind != JsonValueKind.Object) return config;

            if (layoutJson.TryGetProperty("TableColumns", out var tcEl))
                config.TableColumns = tcEl.EnumerateArray().Select(x => x.GetString() ?? "").ToList();

            if (layoutJson.TryGetProperty("ColumnVisibility", out var cvEl))
            {
                foreach (var prop in cvEl.EnumerateObject())
                    config.ColumnVisibility[prop.Name] = prop.Value.GetBoolean();
            }

            if (layoutJson.TryGetProperty("ColumnSorts", out var csEl))
            {
                foreach (var sortItem in csEl.EnumerateArray())
                {
                    config.ColumnSorts.Add(new SortDesc
                    {
                        ColumnId = sortItem.GetProperty("ColumnId").GetString() ?? "",
                        SortOrder = sortItem.GetProperty("SortOrder").GetString() ?? "Asc"
                    });
                }
            }

            if (layoutJson.TryGetProperty("ColumnFilters", out var cfEl))
            {
                foreach (var filterItem in cfEl.EnumerateArray())
                {
                    var filter = new FilterDesc
                    {
                        ColumnId = filterItem.GetProperty("ColumnId").GetString() ?? "",
                        PredicatesOperator = filterItem.TryGetProperty("PredicatesOperator", out var po) ? po.GetString() ?? "AND" : "AND",
                        Predicates = new List<PredicateDesc>()
                    };

                    if (filterItem.TryGetProperty("Predicates", out var predEl))
                    {
                        foreach (var p in predEl.EnumerateArray())
                        {
                            var inputs = new List<object>();
                            if (p.TryGetProperty("Inputs", out var inEl))
                            {
                                foreach (var inp in inEl.EnumerateArray())
                                {
                                    switch (inp.ValueKind)
                                    {
                                        case JsonValueKind.String: inputs.Add(inp.GetString()!); break;
                                        case JsonValueKind.Number: inputs.Add(inp.GetDouble()); break;
                                        case JsonValueKind.True: inputs.Add(true); break;
                                        case JsonValueKind.False: inputs.Add(false); break;
                                    }
                                }
                            }

                            filter.Predicates.Add(new PredicateDesc
                            {
                                PredicateId = p.GetProperty("PredicateId").GetString() ?? "",
                                Inputs = inputs
                            });
                        }
                    }
                    config.ColumnFilters.Add(filter);
                }
            }

            if (layoutJson.TryGetProperty("RowGroupedColumns", out var rgcEl))
                config.RowGroupedColumns = rgcEl.EnumerateArray().Select(x => x.GetString() ?? "").ToList();

            if (layoutJson.TryGetProperty("TableAggregationColumns", out var tacEl))
            {
                foreach (var aggItem in tacEl.EnumerateArray())
                {
                    config.TableAggregationColumns.Add(new AggregationDesc
                    {
                        ColumnId = aggItem.GetProperty("ColumnId").GetString() ?? "",
                        AggFunc = aggItem.GetProperty("AggFunc").GetString() ?? ""
                    });
                }
            }

            if (layoutJson.TryGetProperty("PivotColumns", out var pcEl))
                config.PivotColumns = pcEl.EnumerateArray().Select(x => x.GetString() ?? "").ToList();

            if (layoutJson.TryGetProperty("PivotGroupedColumns", out var pgcEl))
                config.PivotGroupedColumns = pgcEl.EnumerateArray().Select(x => x.GetString() ?? "").ToList();

            if (layoutJson.TryGetProperty("PivotAggregationColumns", out var pacEl))
            {
                foreach (var aggItem in pacEl.EnumerateArray())
                {
                    config.PivotAggregationColumns.Add(new AggregationDesc
                    {
                        ColumnId = aggItem.GetProperty("ColumnId").GetString() ?? "",
                        AggFunc = aggItem.GetProperty("AggFunc").GetString() ?? ""
                    });
                }
            }

            return config;
        }

        private List<IDictionary<string, object>> ApplyFilters(List<IDictionary<string, object>> data, List<FilterDesc> filters)
        {
            return data.Where(row =>
            {
                foreach (var filter in filters)
                {
                    var cellValue = row.ContainsKey(filter.ColumnId) ? row[filter.ColumnId] : null;
                    bool passAll = filter.PredicatesOperator.Equals("AND", StringComparison.OrdinalIgnoreCase);
                    bool passAny = false;

                    foreach (var predicate in filter.Predicates)
                    {
                        bool pass = EvaluatePredicate(cellValue, predicate);
                        if (passAll)
                        {
                            if (!pass) passAll = false;
                        }
                        else
                        {
                            if (pass) passAny = true;
                        }
                    }

                    if (filter.PredicatesOperator.Equals("AND", StringComparison.OrdinalIgnoreCase) && !passAll) return false;
                    if (!filter.PredicatesOperator.Equals("AND", StringComparison.OrdinalIgnoreCase) && !passAny) return false;
                }
                return true;
            }).ToList();
        }

        private List<IDictionary<string, object>> ApplySorting(List<IDictionary<string, object>> data, List<SortDesc> sorts)
        {
            if (!sorts.Any()) return data;

            var query = data.AsEnumerable();
            bool first = true;
            foreach (var sort in sorts)
            {
                if (first)
                {
                    if (sort.SortOrder.Equals("Desc", StringComparison.OrdinalIgnoreCase))
                        query = query.OrderByDescending(x => x.ContainsKey(sort.ColumnId) && x[sort.ColumnId] != null && !string.IsNullOrWhiteSpace(x[sort.ColumnId].ToString()) ? x[sort.ColumnId] : null);
                    else
                        query = query.OrderBy(x => x.ContainsKey(sort.ColumnId) && x[sort.ColumnId] != null && !string.IsNullOrWhiteSpace(x[sort.ColumnId].ToString()) ? x[sort.ColumnId] : null);
                    first = false;
                }
                else
                {
                    var orderedQuery = (IOrderedEnumerable<IDictionary<string, object>>)query;
                    if (sort.SortOrder.Equals("Desc", StringComparison.OrdinalIgnoreCase))
                        query = orderedQuery.ThenByDescending(x => x.ContainsKey(sort.ColumnId) && x[sort.ColumnId] != null && !string.IsNullOrWhiteSpace(x[sort.ColumnId].ToString()) ? x[sort.ColumnId] : null);
                    else
                        query = orderedQuery.ThenBy(x => x.ContainsKey(sort.ColumnId) && x[sort.ColumnId] != null && !string.IsNullOrWhiteSpace(x[sort.ColumnId].ToString()) ? x[sort.ColumnId] : null);
                }
            }
            return query.ToList();
        }

        private void GeneratePivotTable(XLWorkbook workbook, List<IDictionary<string, object>> filteredData, LayoutConfig config, string rawSheetName, int sheetIndex)
        {
            string dataSheetName = SanitizeSheetName(rawSheetName + "_Raw");
            string ptSheetName = SanitizeSheetName(rawSheetName);

            try { workbook.Worksheets.Worksheet(dataSheetName); dataSheetName = SanitizeSheetName(dataSheetName + "_" + sheetIndex); } catch { }
            try { workbook.Worksheets.Worksheet(ptSheetName); ptSheetName = SanitizeSheetName(ptSheetName + "_p" + sheetIndex); } catch { }
            if (dataSheetName == ptSheetName) dataSheetName += "_Raw";

            var dataSheet = workbook.Worksheets.Add(dataSheetName);
            var allKeys = filteredData.Any() ? filteredData.First().Keys.ToList() : new List<string>();

            for (int i = 0; i < allKeys.Count; i++)
            {
                dataSheet.Cell(1, i + 1).Value = allKeys[i];
                dataSheet.Cell(1, i + 1).Style.Font.Bold = true;
            }

            for (int r = 0; r < filteredData.Count; r++)
            {
                var row = filteredData[r];
                for (int c = 0; c < allKeys.Count; c++)
                {
                    var colName = allKeys[c];
                    if (row.TryGetValue(colName, out var v) && v != null && !string.IsNullOrWhiteSpace(v.ToString()))
                        dataSheet.Cell(r + 2, c + 1).Value = XLCellValue.FromObject(v);
                }
            }

            var ptSheet = workbook.Worksheets.Add(ptSheetName);
            if (filteredData.Any())
            {
                var sourceRange = dataSheet.Range(1, 1, filteredData.Count + 1, allKeys.Count);
                var pt = ptSheet.PivotTables.Add("PivotTable1", ptSheet.Cell(1, 1), sourceRange);

                foreach (var grp in config.PivotGroupedColumns) pt.RowLabels.Add(grp);
                foreach (var col in config.PivotColumns) pt.ColumnLabels.Add(col);
                foreach (var agg in config.PivotAggregationColumns)
                {
                    var field = pt.Values.Add(agg.ColumnId);
                    if (agg.AggFunc == "sum") field.SummaryFormula = XLPivotSummary.Sum;
                    else if (agg.AggFunc == "count") field.SummaryFormula = XLPivotSummary.Count;
                    else if (agg.AggFunc == "avg") field.SummaryFormula = XLPivotSummary.Average;
                    else if (agg.AggFunc == "min") field.SummaryFormula = XLPivotSummary.Minimum;
                    else if (agg.AggFunc == "max") field.SummaryFormula = XLPivotSummary.Maximum;
                }

                pt.Theme = XLPivotTableTheme.PivotStyleMedium9;
                ptSheet.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                ptSheet.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            dataSheet.Hide();
            ptSheet.Columns().AdjustToContents();
        }

        private void GenerateGridSheet(XLWorkbook workbook, List<IDictionary<string, object>> filteredData, LayoutConfig config, string rawSheetName, int sheetIndex, List<StyledColumnDesc> styledColumns)
        {
            string rsName = SanitizeSheetName(rawSheetName);
            try { workbook.Worksheets.Worksheet(rsName); rsName = SanitizeSheetName(rsName + "_" + sheetIndex); } catch { }

            var worksheet = workbook.Worksheets.Add(rsName);

            // Determine visible columns
            var tableCols = config.TableColumns;
            if (!tableCols.Any() && filteredData.Any())
            {
                tableCols = filteredData.First().Keys.ToList();
            }
            tableCols.Remove("ag-Grid-AutoColumn");

            var originalVisibleColumns = tableCols.Where(c => !config.ColumnVisibility.TryGetValue(c, out var isVis) || isVis).ToList();
            var visibleColumns = new List<string>();

            if (config.RowGroupedColumns.Any())
            {
                visibleColumns.Add("Group");
            }
            visibleColumns.AddRange(originalVisibleColumns);

            if (!visibleColumns.Any())
            {
                visibleColumns.Add("Data");
            }

            // Headers
            for (int i = 0; i < visibleColumns.Count; i++)
            {
                string headerName = visibleColumns[i];
                var aggDesc = config.TableAggregationColumns.FirstOrDefault(a => a.ColumnId == headerName);
                if (aggDesc != null && !string.IsNullOrWhiteSpace(aggDesc.AggFunc))
                {
                    headerName = $"{aggDesc.AggFunc}({headerName})";
                }

                worksheet.Cell(1, i + 1).Value = headerName;
                worksheet.Cell(1, i + 1).Style.Font.Bold = true;
            }

            int currentRow = 2;

            if (config.RowGroupedColumns.Any())
            {
                WriteGroup(worksheet, filteredData, config.RowGroupedColumns, 0, visibleColumns, config.TableAggregationColumns, config.ColumnSorts, ref currentRow, 1);
                worksheet.Outline.SummaryVLocation = XLOutlineSummaryVLocation.Top;
            }
            else
            {
                foreach (var row in filteredData)
                {
                    for (int c = 0; c < visibleColumns.Count; c++)
                    {
                        var colName = visibleColumns[c];
                        var val = row.ContainsKey(colName) ? row[colName] : null;
                        if (val != null && !string.IsNullOrWhiteSpace(val.ToString()))
                        {
                            worksheet.Cell(currentRow, c + 1).Value = XLCellValue.FromObject(val);
                        }
                    }
                    currentRow++;
                }
            }

            worksheet.Columns().AdjustToContents();
            ApplyConditionalFormatting(worksheet, visibleColumns, styledColumns, currentRow);
        }

        private void ApplyConditionalFormatting(IXLWorksheet worksheet, List<string> visibleColumns, List<StyledColumnDesc> styledColumns, int currentRow)
        {
            foreach (var styledCol in styledColumns)
            {
                int colIndex = visibleColumns.IndexOf(styledCol.ColumnId);
                if (colIndex >= 0)
                {
                    var range = worksheet.Range(2, colIndex + 1, currentRow - 1, colIndex + 1);
                    foreach (var cr in styledCol.GradientRanges)
                    {
                        if (cr.Min.HasValue && cr.Max.HasValue)
                        {
                            range.AddConditionalFormat().WhenBetween(cr.Min.Value, cr.Max.Value)
                                .Fill.SetBackgroundColor(XLColor.FromHtml(cr.Color));
                        }
                        else if (cr.Min.HasValue)
                        {
                            range.AddConditionalFormat().WhenEqualOrGreaterThan(cr.Min.Value)
                                .Fill.SetBackgroundColor(XLColor.FromHtml(cr.Color));
                        }
                        else if (cr.Max.HasValue)
                        {
                            range.AddConditionalFormat().WhenEqualOrLessThan(cr.Max.Value)
                                .Fill.SetBackgroundColor(XLColor.FromHtml(cr.Color));
                        }
                    }

                    // Data bars are currently commented out due to build issues
                    /*foreach (var cr in styledCol.PercentBarRanges)
                    {
                    }*/
                }
            }
        }

        private void WriteGroup(IXLWorksheet worksheet, IEnumerable<IDictionary<string, object>> data, List<string> groupCols, int groupLevel, List<string> visibleCols, List<AggregationDesc> aggs, List<SortDesc> sorts, ref int currentRow, int outlineLevel)
        {
            if (groupLevel >= groupCols.Count)
            {
                foreach (var row in data)
                {
                    for (int c = 1; c < visibleCols.Count; c++)
                    {
                        var colName = visibleCols[c];
                        var val = row.ContainsKey(colName) ? row[colName] : null;
                        if (val != null)
                        {
                            worksheet.Cell(currentRow, c + 1).Value = XLCellValue.FromObject(val);
                        }
                    }
                    if (outlineLevel > 1) worksheet.Row(currentRow).OutlineLevel = outlineLevel - 1;
                    currentRow++;
                }
                return;
            }

            var col = groupCols[groupLevel];
            var grouped = data.GroupBy(r => r.TryGetValue(col, out var val) && val != null && !string.IsNullOrWhiteSpace(val.ToString()) ? val.ToString() : "(Blanks)").ToList();

            var sortDef = sorts.FirstOrDefault(s => s.ColumnId == col);
            bool isDesc = sortDef != null && sortDef.SortOrder.Equals("Desc", StringComparison.OrdinalIgnoreCase);

            var orderedGroups = isDesc 
                ? grouped.OrderByDescending(g => double.TryParse(g.Key, out double d) ? d : (object)(g.Key ?? "")).ToList()
                : grouped.OrderBy(g => double.TryParse(g.Key, out double d) ? d : (object)(g.Key ?? "")).ToList();

            foreach (var grp in orderedGroups)
            {
                string indent = new string(' ', groupLevel * 4);
                string groupHeaderValue = indent + (grp.Key ?? "") + $" ({grp.Count()})";
                worksheet.Cell(currentRow, 1).Value = groupHeaderValue;

                foreach (var agg in aggs)
                {
                    int colIndex = visibleCols.IndexOf(agg.ColumnId);
                    if (colIndex >= 0)
                    {
                        object? aggResult = ComputeAggregate(grp, agg);
                        if (aggResult != null) worksheet.Cell(currentRow, colIndex + 1).Value = XLCellValue.FromObject(aggResult);
                    }
                }

                worksheet.Row(currentRow).Style.Font.Bold = true;
                if (outlineLevel > 1) worksheet.Row(currentRow).OutlineLevel = outlineLevel - 1;
                currentRow++;

                WriteGroup(worksheet, grp.ToList(), groupCols, groupLevel + 1, visibleCols, aggs, sorts, ref currentRow, outlineLevel + 1);
            }
        }

        private object? ComputeAggregate(IEnumerable<IDictionary<string, object>> data, AggregationDesc agg)
        {
            if (agg.AggFunc == "count") return data.Count();
            
            var vals = data.Select(x => x.ContainsKey(agg.ColumnId) && double.TryParse(x[agg.ColumnId]?.ToString(), out double d) ? d : (double?)null)
                .Where(x => x.HasValue)
                .Select(x => x.Value)
                .ToList();

            if (!vals.Any()) return null;

            return agg.AggFunc switch
            {
                "sum" => vals.Sum(),
                "avg" => vals.Average(),
                "min" => vals.Min(),
                "max" => vals.Max(),
                _ => null
            };
        }

        private bool EvaluatePredicate(object? value, PredicateDesc predicate)
        {
            if (predicate.PredicateId == "Blanks")
                return value == null || string.IsNullOrWhiteSpace(value.ToString());
            if (predicate.PredicateId == "NonBlanks")
                return value != null && !string.IsNullOrWhiteSpace(value.ToString());

            if (value == null) return false;
            var strVal = value.ToString() ?? "";
            var input0 = predicate.Inputs.FirstOrDefault();
            var inputStr = input0?.ToString() ?? "";

            return predicate.PredicateId switch
            {
                "Is" or "Equals" => strVal.Equals(inputStr, StringComparison.OrdinalIgnoreCase),
                "IsNot" or "NotEquals" => !strVal.Equals(inputStr, StringComparison.OrdinalIgnoreCase),
                "Contains" => strVal.Contains(inputStr, StringComparison.OrdinalIgnoreCase),
                "NotContains" => !strVal.Contains(inputStr, StringComparison.OrdinalIgnoreCase),
                "StartsWith" => strVal.StartsWith(inputStr, StringComparison.OrdinalIgnoreCase),
                "EndsWith" => strVal.EndsWith(inputStr, StringComparison.OrdinalIgnoreCase),
                "GreaterThan" => double.TryParse(strVal, out double v1) && input0 is double i1 && v1 > i1,
                "GreaterThanOrEqualTo" => double.TryParse(strVal, out double v2) && input0 is double i2 && v2 >= i2,
                "LessThan" => double.TryParse(strVal, out double v3) && input0 is double i3 && v3 < i3,
                "LessThanOrEqualTo" => double.TryParse(strVal, out double v4) && input0 is double i4 && v4 <= i4,
                "Between" or "NotBetween" => EvaluateBetween(strVal, predicate),
                "In" or "NotIn" => EvaluateIn(strVal, predicate),
                _ => true
            };
        }

        private bool EvaluateBetween(string strVal, PredicateDesc predicate)
        {
            if (predicate.Inputs.Count < 2) return false;
            var input0Str = predicate.Inputs[0]?.ToString() ?? "";
            var input1Str = predicate.Inputs[1]?.ToString() ?? "";

            bool isBetween;
            if (double.TryParse(strVal, out double v) && double.TryParse(input0Str, out double min) && double.TryParse(input1Str, out double max))
                isBetween = v >= min && v <= max;
            else if (DateTime.TryParse(strVal, out DateTime dt) && DateTime.TryParse(input0Str, out DateTime dtMin) && DateTime.TryParse(input1Str, out DateTime dtMax))
                isBetween = dt >= dtMin && dt <= dtMax;
            else
                return false;

            return predicate.PredicateId == "Between" ? isBetween : !isBetween;
        }

        private bool EvaluateIn(string strVal, PredicateDesc predicate)
        {
            bool isIn = predicate.Inputs.Any(inp => (inp?.ToString() ?? "").Equals(strVal, StringComparison.OrdinalIgnoreCase));
            return predicate.PredicateId == "In" ? isIn : !isIn;
        }
    }

    public class LayoutConfig
    {
        public List<string> TableColumns { get; set; } = new();
        public Dictionary<string, bool> ColumnVisibility { get; set; } = new();
        public List<SortDesc> ColumnSorts { get; set; } = new();
        public List<FilterDesc> ColumnFilters { get; set; } = new();
        public List<string> RowGroupedColumns { get; set; } = new();
        public List<AggregationDesc> TableAggregationColumns { get; set; } = new();
        public List<string> PivotColumns { get; set; } = new();
        public List<string> PivotGroupedColumns { get; set; } = new();
        public List<AggregationDesc> PivotAggregationColumns { get; set; } = new();
    }

    public class SortDesc { public string ColumnId { get; set; } = ""; public string SortOrder { get; set; } = "Asc"; }
    public class FilterDesc { public string ColumnId { get; set; } = ""; public string PredicatesOperator { get; set; } = "AND"; public List<PredicateDesc> Predicates { get; set; } = new(); }
    public class PredicateDesc { public string PredicateId { get; set; } = ""; public List<object> Inputs { get; set; } = new(); }
    public class AggregationDesc { public string ColumnId { get; set; } = ""; public string AggFunc { get; set; } = ""; }
    public class StyledColumnDesc { public string ColumnId { get; set; } = ""; public List<CellRangeDesc> GradientRanges { get; set; } = new(); public List<CellRangeDesc> PercentBarRanges { get; set; } = new(); }
    public class CellRangeDesc { public double? Min { get; set; } public double? Max { get; set; } public string Color { get; set; } = ""; }
}
