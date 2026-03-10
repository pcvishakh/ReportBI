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
            List<string> tableColumns = new();
            Dictionary<string, bool> columnVisibility = new();
            List<SortDesc> columnSorts = new();
            List<FilterDesc> columnFilters = new();

            List<string> rowGroupedColumns = new();
            List<AggregationDesc> tableAggregationColumns = new();
            List<string> pivotColumns = new();
            List<string> pivotGroupedColumns = new();
            List<AggregationDesc> pivotAggregationColumns = new();

            if (!string.IsNullOrWhiteSpace(gridStateJson))
            {
                try
                {
                    var doc = JsonDocument.Parse(gridStateJson);
                    if (doc.RootElement.TryGetProperty("Layout", out var layoutEl) &&
                        layoutEl.TryGetProperty("Layouts", out var layoutsEl) &&
                        layoutsEl.GetArrayLength() > 0)
                    {
                        var defaultLayout = layoutsEl[0]; // Assume first/default
                        if (defaultLayout.TryGetProperty("TableColumns", out var tcEl))
                        {
                            tableColumns = tcEl.EnumerateArray().Select(x => x.GetString() ?? "").ToList();
                        }

                        if (defaultLayout.TryGetProperty("ColumnVisibility", out var cvEl))
                        {
                            foreach (var prop in cvEl.EnumerateObject())
                            {
                                columnVisibility[prop.Name] = prop.Value.GetBoolean();
                            }
                        }

                        if (defaultLayout.TryGetProperty("ColumnSorts", out var csEl))
                        {
                            foreach (var sortItem in csEl.EnumerateArray())
                            {
                                columnSorts.Add(new SortDesc
                                {
                                    ColumnId = sortItem.GetProperty("ColumnId").GetString() ?? "",
                                    SortOrder = sortItem.GetProperty("SortOrder").GetString() ?? "Asc"
                                });
                            }
                        }

                        if (defaultLayout.TryGetProperty("ColumnFilters", out var cfEl))
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
                                                    case JsonValueKind.String:
                                                        inputs.Add(inp.GetString()!);
                                                        break;
                                                    case JsonValueKind.Number:
                                                        inputs.Add(inp.GetDouble());
                                                        break;
                                                    case JsonValueKind.True:
                                                        inputs.Add(true);
                                                        break;
                                                    case JsonValueKind.False:
                                                        inputs.Add(false);
                                                        break;
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

                                columnFilters.Add(filter);
                            }
                        }

                        if (defaultLayout.TryGetProperty("RowGroupedColumns", out var rgcEl))
                        {
                            rowGroupedColumns = rgcEl.EnumerateArray().Select(x => x.GetString() ?? "").ToList();
                        }

                        if (defaultLayout.TryGetProperty("TableAggregationColumns", out var tacEl))
                        {
                            foreach (var aggItem in tacEl.EnumerateArray())
                            {
                                tableAggregationColumns.Add(new AggregationDesc
                                {
                                    ColumnId = aggItem.GetProperty("ColumnId").GetString() ?? "",
                                    AggFunc = aggItem.GetProperty("AggFunc").GetString() ?? ""
                                });
                            }
                        }

                        if (defaultLayout.TryGetProperty("PivotColumns", out var pcEl))
                            pivotColumns = pcEl.EnumerateArray().Select(x => x.GetString() ?? "").ToList();

                        if (defaultLayout.TryGetProperty("PivotGroupedColumns", out var pgcEl))
                            pivotGroupedColumns = pgcEl.EnumerateArray().Select(x => x.GetString() ?? "").ToList();

                        if (defaultLayout.TryGetProperty("PivotAggregationColumns", out var pacEl))
                        {
                            foreach (var aggItem in pacEl.EnumerateArray())
                            {
                                pivotAggregationColumns.Add(new AggregationDesc
                                {
                                    ColumnId = aggItem.GetProperty("ColumnId").GetString() ?? "",
                                    AggFunc = aggItem.GetProperty("AggFunc").GetString() ?? ""
                                });
                            }
                        }
                    }
                }
                catch { }
            }

            // Apply Filters
            var filteredData = dataList.Where(row =>
            {
                foreach (var filter in columnFilters)
                {
                    var cellValue = row.ContainsKey(filter.ColumnId) ? row[filter.ColumnId] : null;
                    bool passAll = filter.PredicatesOperator.ToUpper() == "AND";
                    bool passAny = false;

                    foreach (var predicate in filter.Predicates)
                    {
                        bool pass = EvaluatePredicate(cellValue, predicate);
                        if (filter.PredicatesOperator.ToUpper() == "AND")
                        {
                            if (!pass) passAll = false;
                        }
                        else
                        {
                            if (pass) passAny = true;
                        }
                    }

                    if (filter.PredicatesOperator.ToUpper() == "AND" && !passAll) return false;
                    if (filter.PredicatesOperator.ToUpper() != "AND" && !passAny) return false;
                }
                return true;
            }).ToList();

            // Apply Sorting
            if (columnSorts.Any())
            {
                var query = filteredData.AsEnumerable();
                bool first = true;
                foreach (var sort in columnSorts)
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
                filteredData = query.ToList();
            }

            // Determine visible columns
            if (!tableColumns.Any() && filteredData.Any())
            {
                tableColumns = filteredData.First().Keys.ToList();
            }

            tableColumns.Remove("ag-Grid-AutoColumn");
            var originalVisibleColumns = tableColumns.Where(c => !columnVisibility.TryGetValue(c, out var isVis) || isVis).ToList();
            var visibleColumns = new List<string>();

            if (rowGroupedColumns.Any())
            {
                visibleColumns.Add("Group");
            }
            visibleColumns.AddRange(originalVisibleColumns);

            if (!visibleColumns.Any()) 
            {
                visibleColumns.Add("Data");
            }

            // Build Excel
            using var workbook = new XLWorkbook();

            if (pivotColumns.Any() || pivotGroupedColumns.Any())
            {
                var dataSheet = workbook.Worksheets.Add("RawData");
                var allKeys = filteredData.Any() ? filteredData.First().Keys.ToList() : new List<string>();
                
                for (int i = 0; i < allKeys.Count; i++) {
                    dataSheet.Cell(1, i + 1).Value = allKeys[i];
                    dataSheet.Cell(1, i + 1).Style.Font.Bold = true;
                }
                
                for (int r = 0; r < filteredData.Count; r++) {
                    var row = filteredData[r];
                    for (int c = 0; c < allKeys.Count; c++) {
                        var colName = allKeys[c];
                        if (row.TryGetValue(colName, out var v) && v != null && !string.IsNullOrWhiteSpace(v.ToString()))
                            dataSheet.Cell(r + 2, c + 1).Value = XLCellValue.FromObject(v);
                    }
                }

                var ptSheet = workbook.Worksheets.Add("Pivot Export");
                if (filteredData.Any()) 
                {
                    var sourceRange = dataSheet.Range(1, 1, filteredData.Count + 1, allKeys.Count);
                    var pt = ptSheet.PivotTables.Add("PivotTable1", ptSheet.Cell(1, 1), sourceRange);

                    foreach (var grp in pivotGroupedColumns) pt.RowLabels.Add(grp);
                    foreach (var col in pivotColumns) pt.ColumnLabels.Add(col);
                    foreach (var agg in pivotAggregationColumns) 
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
            else
            {
                var worksheet = workbook.Worksheets.Add("Export");

                // Headers
                for (int i = 0; i < visibleColumns.Count; i++)
                {
                    string headerName = visibleColumns[i];
                    var aggDesc = tableAggregationColumns.FirstOrDefault(a => a.ColumnId == headerName);
                    if (aggDesc != null && !string.IsNullOrWhiteSpace(aggDesc.AggFunc))
                    {
                        headerName = $"{aggDesc.AggFunc}({headerName})";
                    }

                    worksheet.Cell(1, i + 1).Value = headerName;
                    worksheet.Cell(1, i + 1).Style.Font.Bold = true;
                }

                int currentRow = 2;

                if (rowGroupedColumns.Any())
                {
                    WriteGroup(worksheet, filteredData, rowGroupedColumns, 0, visibleColumns, tableAggregationColumns, columnSorts, ref currentRow, 1);
                    
                    // Configure outline settings so groups are collapsible
                    worksheet.Outline.SummaryVLocation = XLOutlineSummaryVLocation.Top;
                }
                else
                {
                    // Flat Data
                    for (int r = 0; r < filteredData.Count; r++)
                    {
                        var row = filteredData[r];
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
            }

            using var ms = new MemoryStream();
            workbook.SaveAs(ms);
            return ms.ToArray();
        }

        private void WriteGroup(IXLWorksheet worksheet, IEnumerable<IDictionary<string, object>> data, List<string> groupCols, int groupLevel, List<string> visibleCols, List<AggregationDesc> aggs, List<SortDesc> sorts, ref int currentRow, int outlineLevel)
        {
            if (groupLevel >= groupCols.Count)
            {
                foreach (var row in data)
                {
                    // Skip the 'Group' columns when rendering leaf row data
                    for (int c = 1; c < visibleCols.Count; c++)
                    {
                        var colName = visibleCols[c];
                        var val = row.ContainsKey(colName) ? row[colName] : null;
                        if (val != null)
                        {
                            worksheet.Cell(currentRow, c + 1).Value = XLCellValue.FromObject(val);
                        }
                    }
                    if (outlineLevel > 1) 
                    {
                        worksheet.Row(currentRow).OutlineLevel = outlineLevel - 1;
                    }
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
                string groupHeaderValue = indent + (grp.Key ?? "");
                groupHeaderValue += $" ({grp.Count()})";

                worksheet.Cell(currentRow, 1).Value = groupHeaderValue;

                foreach (var agg in aggs)
                {
                    // Find the index of the aggregated column in visible columns
                    int colIndex = visibleCols.IndexOf(agg.ColumnId);
                    
                    if (colIndex >= 0)
                    {
                        object? aggResult = null;
                        
                        if (agg.AggFunc == "count")
                        {
                            aggResult = grp.Count();
                        }
                        else if (agg.AggFunc == "sum" || agg.AggFunc == "avg" || agg.AggFunc == "min" || agg.AggFunc == "max")
                        {
                            var vals = grp.Select(x => x.ContainsKey(agg.ColumnId) && double.TryParse(x[agg.ColumnId]?.ToString(), out double d) ? d : (double?)null)
                                .Where(x => x.HasValue)
                                .Select(x => x.Value)
                                .ToList();
                            
                            if (vals.Any())
                            {
                                if (agg.AggFunc == "sum") aggResult = vals.Sum();
                                if (agg.AggFunc == "avg") aggResult = vals.Average();
                                if (agg.AggFunc == "min") aggResult = vals.Min();
                                if (agg.AggFunc == "max") aggResult = vals.Max();
                            }
                        }

                        if (aggResult != null)
                        {
                            // Output the raw aggregate result dynamically into the corresponding column of the Group row
                            worksheet.Cell(currentRow, colIndex + 1).Value = XLCellValue.FromObject(aggResult);
                        }
                    }
                }

                worksheet.Row(currentRow).Style.Font.Bold = true;
                if (outlineLevel > 1) 
                {
                    worksheet.Row(currentRow).OutlineLevel = outlineLevel - 1;
                }
                currentRow++;

                WriteGroup(worksheet, grp.ToList(), groupCols, groupLevel + 1, visibleCols, aggs, sorts, ref currentRow, outlineLevel + 1);
            }
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

            switch (predicate.PredicateId)
            {
                case "Is":
                case "Equals":
                    return strVal.Equals(inputStr, StringComparison.OrdinalIgnoreCase);
                case "IsNot":
                case "NotEquals":
                    return !strVal.Equals(inputStr, StringComparison.OrdinalIgnoreCase);
                case "Contains":
                    return strVal.Contains(inputStr, StringComparison.OrdinalIgnoreCase);
                case "NotContains":
                    return !strVal.Contains(inputStr, StringComparison.OrdinalIgnoreCase);
                case "StartsWith":
                    return strVal.StartsWith(inputStr, StringComparison.OrdinalIgnoreCase);
                case "EndsWith":
                    return strVal.EndsWith(inputStr, StringComparison.OrdinalIgnoreCase);
                case "GreaterThan":
                    if (double.TryParse(strVal, out double v1) && input0 is double i1)
                        return v1 > i1;
                    return false;
                case "GreaterThanOrEqualTo":
                    if (double.TryParse(strVal, out double v2) && input0 is double i2)
                        return v2 >= i2;
                    return false;
                case "LessThan":
                    if (double.TryParse(strVal, out double v3) && input0 is double i3)
                        return v3 < i3;
                    return false;
                case "LessThanOrEqualTo":
                    if (double.TryParse(strVal, out double v4) && input0 is double i4)
                        return v4 <= i4;
                    return false;
                case "Between":
                case "NotBetween":
                    if (predicate.Inputs.Count > 1)
                    {
                        var input1Str = predicate.Inputs[1]?.ToString() ?? "";
                        
                        // Try numeric between
                        if (double.TryParse(strVal, out double v5) && 
                            double.TryParse(inputStr, out double minNum) && 
                            double.TryParse(input1Str, out double maxNum))
                        {
                            bool isBetween = v5 >= minNum && v5 <= maxNum;
                            return predicate.PredicateId == "Between" ? isBetween : !isBetween;
                        }

                        // Try datetime between
                        if (DateTime.TryParse(strVal, out DateTime dtVal) && 
                            DateTime.TryParse(inputStr, out DateTime dtMin) && 
                            DateTime.TryParse(input1Str, out DateTime dtMax))
                        {
                            bool isBetween = dtVal >= dtMin && dtVal <= dtMax;
                            return predicate.PredicateId == "Between" ? isBetween : !isBetween;
                        }
                    }
                    return false;
                case "In":
                case "NotIn":
                    if (predicate.Inputs.Count > 0)
                    {
                        bool isIn = predicate.Inputs.Any(inp => 
                            (inp?.ToString() ?? "").Equals(strVal, StringComparison.OrdinalIgnoreCase)
                        );
                        return predicate.PredicateId == "In" ? isIn : !isIn;
                    }
                    return false;
                default:
                    return true;
            }
        }
    }

    public class SortDesc
    {
        public string ColumnId { get; set; } = "";
        public string SortOrder { get; set; } = "Asc";
    }

    public class FilterDesc
    {
        public string ColumnId { get; set; } = "";
        public string PredicatesOperator { get; set; } = "AND";
        public List<PredicateDesc> Predicates { get; set; } = new();
    }

    public class PredicateDesc
    {
        public string PredicateId { get; set; } = "";
        public List<object> Inputs { get; set; } = new();
    }

    public class AggregationDesc
    {
        public string ColumnId { get; set; } = "";
        public string AggFunc { get; set; } = "";
    }
}
