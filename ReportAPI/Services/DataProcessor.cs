using ReportAPI.Models;

namespace ReportAPI.Services
{
    public interface IDataProcessor
    {
        List<IDictionary<string, object>> ProcessData(List<IDictionary<string, object>> data, LayoutConfig config);
        object? ComputeAggregate(IEnumerable<IDictionary<string, object>> data, AggregationDesc agg);
        bool IsJson(object? val);
    }

    public class DataProcessor : IDataProcessor
    {
        private const string BlankLabel = "(Blank)";

        public List<IDictionary<string, object>> ProcessData(List<IDictionary<string, object>> data, LayoutConfig config)
        {
            var filteredData = ApplyFilters(data, config.ColumnFilters);
            return ApplySorting(filteredData, config.ColumnSorts);
        }

        private List<IDictionary<string, object>> ApplyFilters(List<IDictionary<string, object>> data, List<FilterDesc> filters)
        {
            if (filters == null || !filters.Any()) return data;

            return data.Where(row =>
            {
                foreach (var filter in filters)
                {
                    var cellValue = row.TryGetValue(filter.ColumnId, out var v) ? v : null;
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

        private List<IDictionary<string, object>> ApplySorting(List<IDictionary<string, object>> data, List<SortDesc> sorts)
        {
            if (sorts == null || !sorts.Any()) return data;

            IOrderedEnumerable<IDictionary<string, object>>? query = null;

            foreach (var sort in sorts)
            {
                Func<IDictionary<string, object>, object?> keySelector = x => x.TryGetValue(sort.ColumnId, out var v) && v != null && !string.IsNullOrWhiteSpace(v.ToString()) ? v : null;
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

        public object? ComputeAggregate(IEnumerable<IDictionary<string, object>> data, AggregationDesc agg)
        {
            if (agg == null) return null;
            if (agg.AggFunc == "count") return $"({data.Count()})";

            var vals = data.Select(x => x.TryGetValue(agg.ColumnId, out var val) && double.TryParse(val?.ToString(), out double d) ? d : (double?)null)
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

        public bool IsJson(object? val)
        {
            if (val == null) return false;
            var s = val.ToString()?.Trim();
            if (string.IsNullOrEmpty(s)) return false;
            return (s.StartsWith("[") && s.EndsWith("]")) || (s.StartsWith("{") && s.EndsWith("}"));
        }
    }
}
