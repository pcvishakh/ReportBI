using ReportAPI.Models;

namespace ReportAPI.Services
{
    public interface IDataProcessor
    {
        List<IDictionary<string, object>> ProcessData(List<IDictionary<string, object>> data, LayoutConfig config, IEnumerable<ColumnDefinition>? columnDefinitions = null);
        object? ComputeAggregate(IEnumerable<IDictionary<string, object>> data, AggregationDesc agg, ColumnDefinition? colDef = null);
        bool IsJson(object? val);
    }

    public class DataProcessor : IDataProcessor
    {
        private const string BlankLabel = "(Blank)";

        public List<IDictionary<string, object>> ProcessData(List<IDictionary<string, object>> data, LayoutConfig config, IEnumerable<ColumnDefinition>? columnDefinitions = null)
        {
            var filteredData = ApplyFilters(data, config.ColumnFilters, columnDefinitions);
            return ApplySorting(filteredData, config.ColumnSorts, columnDefinitions);
        }

        private List<IDictionary<string, object>> ApplyFilters(List<IDictionary<string, object>> data, List<FilterDesc> filters, IEnumerable<ColumnDefinition>? columnDefinitions)
        {
            if (filters == null || !filters.Any()) return data;

            return data.Where(row =>
            {
                foreach (var filter in filters)
                {
                    var colDef = columnDefinitions?.FirstOrDefault(c => c.Field == filter.ColumnId);
                    var cellValue = row.TryGetValue(filter.ColumnId, out var v) ? v : null;
                    bool passAll = filter.PredicatesOperator.Equals("AND", StringComparison.OrdinalIgnoreCase);
                    bool passAny = false;

                    foreach (var predicate in filter.Predicates)
                    {
                        bool pass = EvaluatePredicate(cellValue, predicate, colDef);
                        if (passAll)
                        {
                            if (!pass) passAll = false;
                        }
                        else
                        {
                            if (pass) passAny = true;
                        }
                    }

                    if (filter.PredicatesOperator.Equals("AND", StringComparison.OrdinalIgnoreCase))
                    {
                        if (!passAll) return false;
                    }
                    else
                    {
                        if (!passAny) return false;
                    }
                }
                return true;
            }).ToList();
        }

        private List<IDictionary<string, object>> ApplySorting(List<IDictionary<string, object>> data, List<SortDesc> sorts, IEnumerable<ColumnDefinition>? columnDefinitions)
        {
            if (sorts == null || !sorts.Any()) return data;

            IOrderedEnumerable<IDictionary<string, object>>? query = null;

            foreach (var sort in sorts)
            {
                var colDef = columnDefinitions?.FirstOrDefault(c => c.Field == sort.ColumnId);
                Func<IDictionary<string, object>, object?> keySelector = x =>
                {
                    if (x.TryGetValue(sort.ColumnId, out var v) && v != null && !string.IsNullOrWhiteSpace(v.ToString()))
                    {
                        if (colDef?.DataType == "number" && double.TryParse(v.ToString(), out double d)) return d;
                        if (colDef?.DataType == "date" && DateTime.TryParse(v.ToString(), out DateTime dt)) return dt;
                        return v;
                    }
                    return null;
                };
                
                bool isDesc = sort.SortOrder.Equals("Desc", StringComparison.OrdinalIgnoreCase);

                if (query == null)
                {
                    query = isDesc ? data.OrderByDescending(keySelector) : data.OrderBy(keySelector);
                }
                else
                {
                    query = isDesc ? query.ThenByDescending(keySelector) : query.ThenBy(keySelector);
                }
            }

            return query?.ToList() ?? data;
        }

        public object? ComputeAggregate(IEnumerable<IDictionary<string, object>> data, AggregationDesc agg, ColumnDefinition? colDef = null)
        {
            if (agg == null) return null;
            if (string.Equals(agg.AggFunc, "count", StringComparison.OrdinalIgnoreCase)) return (long)data.Count();

            var vals = data.Select(x =>
                {
                    if (x.TryGetValue(agg.ColumnId, out var val) && val != null)
                    {
                        if (colDef?.DataType == "number" && double.TryParse(val.ToString(), out double d)) return d;
                        if (double.TryParse(val.ToString(), out double dFallback)) return dFallback;
                    }
                    return (double?)null;
                })
                .Where(x => x.HasValue)
                .Select(x => x!.Value)
                .ToList();

            if (!vals.Any()) return null;

            return agg.AggFunc.ToLower() switch
            {
                "sum" => vals.Sum(),
                "avg" => vals.Average(),
                "min" => vals.Min(),
                "max" => vals.Max(),
                _ => null
            };
        }

        private bool EvaluatePredicate(object? value, PredicateDesc predicate, ColumnDefinition? colDef)
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
                "GreaterThan" => CompareValues(strVal, input0, colDef) > 0,
                "GreaterThanOrEqualTo" => CompareValues(strVal, input0, colDef) >= 0,
                "LessThan" => CompareValues(strVal, input0, colDef) < 0,
                "LessThanOrEqualTo" => CompareValues(strVal, input0, colDef) <= 0,
                "Between" or "NotBetween" => EvaluateBetween(strVal, predicate, colDef),
                "In" or "NotIn" => EvaluateIn(strVal, predicate),
                _ => true
            };
        }

        private int CompareValues(string strVal, object? input, ColumnDefinition? colDef)
        {
            if (input == null) return 0;
            var inputStr = input.ToString() ?? "";

            if (colDef?.DataType == "number" || double.TryParse(strVal, out _) || double.TryParse(inputStr, out _))
            {
                if (double.TryParse(strVal, out double v1) && double.TryParse(inputStr, out double v2))
                    return v1.CompareTo(v2);
            }
            if (colDef?.DataType == "date" || DateTime.TryParse(strVal, out _) || DateTime.TryParse(inputStr, out _))
            {
                if (DateTime.TryParse(strVal, out DateTime d1) && DateTime.TryParse(inputStr, out DateTime d2))
                    return d1.CompareTo(d2);
            }

            return string.Compare(strVal, inputStr, StringComparison.OrdinalIgnoreCase);
        }

        private bool EvaluateBetween(string strVal, PredicateDesc predicate, ColumnDefinition? colDef)
        {
            if (predicate.Inputs.Count < 2) return false;
            var input0Str = predicate.Inputs[0]?.ToString() ?? "";
            var input1Str = predicate.Inputs[1]?.ToString() ?? "";

            bool isBetween;
            if ((colDef?.DataType == "number" || double.TryParse(strVal, out _)) && double.TryParse(input0Str, out double min) && double.TryParse(input1Str, out double max))
            {
                if (double.TryParse(strVal, out double v))
                    isBetween = v >= min && v <= max;
                else
                    return false;
            }
            else if ((colDef?.DataType == "date" || DateTime.TryParse(strVal, out _)) && DateTime.TryParse(input0Str, out DateTime dtMin) && DateTime.TryParse(input1Str, out DateTime dtMax))
            {
                if (DateTime.TryParse(strVal, out DateTime dt))
                    isBetween = dt >= dtMin && dt <= dtMax;
                else
                    return false;
            }
            else
            {
                isBetween = string.Compare(strVal, input0Str, StringComparison.OrdinalIgnoreCase) >= 0 &&
                            string.Compare(strVal, input1Str, StringComparison.OrdinalIgnoreCase) <= 0;
            }

            return predicate.PredicateId == "Between" ? isBetween : !isBetween;
        }

        private bool EvaluateIn(string strVal, PredicateDesc predicate)
        {
            bool isIn = predicate.Inputs.Any(inp => (inp?.ToString() ?? "").Equals(strVal, StringComparison.OrdinalIgnoreCase));
            return predicate.PredicateId == "In" ? isIn : !isIn;
        }

        public bool IsJson(object? val)
        {
            if (val == null) return false;
            var s = val.ToString()?.Trim();
            if (string.IsNullOrEmpty(s)) return false;
            return (s.StartsWith("[") && s.EndsWith("]")) || (s.StartsWith("{") && s.EndsWith("}"));
        }
    }
}
