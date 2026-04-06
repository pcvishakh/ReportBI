using System.Text.Json;

namespace ReportAPI.Models
{
    public class LayoutConfig
    {
        public string Name { get; set; } = string.Empty;
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
