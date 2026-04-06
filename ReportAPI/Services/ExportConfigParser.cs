using System.Text.Json;
using ReportAPI.Models;

namespace ReportAPI.Services
{
    public interface IExportConfigParser
    {
        List<LayoutConfig> ParseGridState(string? gridStateJson);
    }

    public class ExportConfigParser : IExportConfigParser
    {
        public List<LayoutConfig> ParseGridState(string? gridStateJson)
        {
            var layouts = new List<LayoutConfig>();

            if (string.IsNullOrWhiteSpace(gridStateJson))
                return layouts;

            try
            {
                using var doc = JsonDocument.Parse(gridStateJson);
                var root = doc.RootElement;

                if (root.TryGetProperty("Layout", out var layoutEl) &&
                    layoutEl.TryGetProperty("Layouts", out var layoutsEl) &&
                    layoutsEl.ValueKind == JsonValueKind.Array)
                {
                    foreach (var layoutItem in layoutsEl.EnumerateArray())
                    {
                        var config = ParseLayoutConfig(layoutItem);
                        layouts.Add(config);
                    }
                }
            }
            catch 
            {
                // Silently fail as in the original code
            }

            return layouts;
        }

        private LayoutConfig ParseLayoutConfig(JsonElement layoutJson)
        {
            var config = new LayoutConfig();
            if (layoutJson.ValueKind != JsonValueKind.Object) return config;

            if (layoutJson.TryGetProperty("Name", out var nameProp))
                config.Name = nameProp.GetString() ?? "";

            if (layoutJson.TryGetProperty("TableColumns", out var tcEl) && tcEl.ValueKind == JsonValueKind.Array)
                config.TableColumns = tcEl.EnumerateArray().Select(x => x.GetString() ?? "").ToList();

            if (layoutJson.TryGetProperty("ColumnVisibility", out var cvEl) && cvEl.ValueKind == JsonValueKind.Object)
            {
                foreach (var prop in cvEl.EnumerateObject())
                    config.ColumnVisibility[prop.Name] = prop.Value.ValueKind == JsonValueKind.True;
            }

            if (layoutJson.TryGetProperty("ColumnSorts", out var csEl) && csEl.ValueKind == JsonValueKind.Array)
            {
                foreach (var sortItem in csEl.EnumerateArray())
                {
                    config.ColumnSorts.Add(new SortDesc
                    {
                        ColumnId = sortItem.TryGetProperty("ColumnId", out var ci) ? ci.GetString() ?? "" : "",
                        SortOrder = sortItem.TryGetProperty("SortOrder", out var so) ? so.GetString() ?? "Asc" : "Asc"
                    });
                }
            }

            if (layoutJson.TryGetProperty("ColumnFilters", out var cfEl) && cfEl.ValueKind == JsonValueKind.Array)
            {
                foreach (var filterItem in cfEl.EnumerateArray())
                {
                    var filter = new FilterDesc
                    {
                        ColumnId = filterItem.TryGetProperty("ColumnId", out var ci) ? ci.GetString() ?? "" : "",
                        PredicatesOperator = filterItem.TryGetProperty("PredicatesOperator", out var po) ? po.GetString() ?? "AND" : "AND",
                        Predicates = new List<PredicateDesc>()
                    };

                    if (filterItem.TryGetProperty("Predicates", out var predEl) && predEl.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var p in predEl.EnumerateArray())
                        {
                            var inputs = new List<object>();
                            if (p.TryGetProperty("Inputs", out var inEl) && inEl.ValueKind == JsonValueKind.Array)
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
                                PredicateId = p.TryGetProperty("PredicateId", out var pid) ? pid.GetString() ?? "" : "",
                                Inputs = inputs
                            });
                        }
                    }
                    config.ColumnFilters.Add(filter);
                }
            }

            if (layoutJson.TryGetProperty("RowGroupedColumns", out var rgcEl) && rgcEl.ValueKind == JsonValueKind.Array)
                config.RowGroupedColumns = rgcEl.EnumerateArray().Select(x => x.GetString() ?? "").ToList();

            if (layoutJson.TryGetProperty("TableAggregationColumns", out var tacEl) && tacEl.ValueKind == JsonValueKind.Array)
            {
                foreach (var aggItem in tacEl.EnumerateArray())
                {
                    config.TableAggregationColumns.Add(new AggregationDesc
                    {
                        ColumnId = aggItem.TryGetProperty("ColumnId", out var ci) ? ci.GetString() ?? "" : "",
                        AggFunc = aggItem.TryGetProperty("AggFunc", out var af) ? af.GetString() ?? "" : ""
                    });
                }
            }

            if (layoutJson.TryGetProperty("PivotColumns", out var pcEl) && pcEl.ValueKind == JsonValueKind.Array)
                config.PivotColumns = pcEl.EnumerateArray().Select(x => x.GetString() ?? "").ToList();

            if (layoutJson.TryGetProperty("PivotGroupedColumns", out var pgcEl) && pgcEl.ValueKind == JsonValueKind.Array)
                config.PivotGroupedColumns = pgcEl.EnumerateArray().Select(x => x.GetString() ?? "").ToList();

            if (layoutJson.TryGetProperty("PivotAggregationColumns", out var pacEl) && pacEl.ValueKind == JsonValueKind.Array)
            {
                foreach (var aggItem in pacEl.EnumerateArray())
                {
                    config.PivotAggregationColumns.Add(new AggregationDesc
                    {
                        ColumnId = aggItem.TryGetProperty("ColumnId", out var ci) ? ci.GetString() ?? "" : "",
                        AggFunc = aggItem.TryGetProperty("AggFunc", out var af) ? af.GetString() ?? "" : ""
                    });
                }
            }

            return config;
        }
    }
}
