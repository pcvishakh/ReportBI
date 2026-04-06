
using System;
using System.Collections.Generic;
using System.Linq;

public class AggregationDesc {
    public string ColumnId { get; set; }
    public string AggFunc { get; set; }
}

public class DataProcessor {
    public object? ComputeAggregate(IEnumerable<IDictionary<string, object>> data, AggregationDesc agg)
    {
        if (agg == null) return null;
        if (agg.AggFunc == "count") return $"{data.Count()}";

        var vals = data.Select(x =>
            {
                if (x.TryGetValue(agg.ColumnId, out var val) && val != null)
                {
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
}

class Program {
    static void Main() {
        var dp = new DataProcessor();
        var data = new List<IDictionary<string, object>> {
            new Dictionary<string, object> { {"Val", 10} },
            new Dictionary<string, object> { {"Val", null} },
            new Dictionary<string, object> { {"Val", 20} }
        };

        var agg1 = new AggregationDesc { ColumnId = "Val", AggFunc = "count" };
        Console.WriteLine($"count (lowercase): {dp.ComputeAggregate(data, agg1)}");

        var agg2 = new AggregationDesc { ColumnId = "Val", AggFunc = "Count" };
        Console.WriteLine($"Count (capital): {dp.ComputeAggregate(data, agg2) ?? "null"}");
    }
}
