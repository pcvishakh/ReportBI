using System.Text.Json;
using ReportAPI.Models;

namespace ReportAPI.Services
{
    public interface IExportConfigParser
    {
        (List<StyledColumnDesc> StyledColumns, List<LayoutConfig> Layouts) ParseGridState(string? gridStateJson);
    }

    public class ExportConfigParser : IExportConfigParser
    {
        public (List<StyledColumnDesc> StyledColumns, List<LayoutConfig> Layouts) ParseGridState(string? gridStateJson)
        {
            var styledColumns = new List<StyledColumnDesc>();
            var layouts = new List<LayoutConfig>();

            if (string.IsNullOrWhiteSpace(gridStateJson))
                return (styledColumns, layouts);

            try
            {
                using var doc = JsonDocument.Parse(gridStateJson);
                var root = doc.RootElement;

                styledColumns = ParseStyledColumns(root);

                if (root.TryGetProperty("Layout", out var layoutEl) &&
                    layoutEl.TryGetProperty("Layouts", out var layoutsEl) &&
                    layoutsEl.GetArrayLength() > 0)
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

            return (styledColumns, layouts);
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
                Min = crItem.TryGetProperty("Min", out var minV) && minV.ValueKind == JsonValueKind.Number ? minV.GetDouble() : (double?)null,
                Max = crItem.TryGetProperty("Max", out var maxV) && maxV.ValueKind == JsonValueKind.Number ? maxV.GetDouble() : (double?)null,
                Color = crItem.TryGetProperty("Color", out var colV) ? colV.GetString() ?? "" : ""
            };
        }

        private LayoutConfig ParseLayoutConfig(JsonElement layoutJson)
        {
            var config = new LayoutConfig();
            if (layoutJson.ValueKind != JsonValueKind.Object) return config;

            if (layoutJson.TryGetProperty("Name", out var nameProp))
                config.Name = nameProp.GetString() ?? "";

            if (layoutJson.TryGetProperty("TableColumns", out var tcEl))
                config.TableColumns = tcEl.EnumerateArray().Select(x => x.GetString() ?? "").ToList();

            if (layoutJson.TryGetProperty("ColumnVisibility", out var cvEl))
            {
                foreach (var prop in cvEl.EnumerateObject())
                    config.ColumnVisibility[prop.Name] = prop.Value.ValueKind == JsonValueKind.True;
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
    }
}
