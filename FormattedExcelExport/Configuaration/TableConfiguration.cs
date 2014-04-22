using System.Collections.Generic;
using FormattedExcelExport.Infrastructure;


namespace FormattedExcelExport.Configuaration {
	public class TableConfiguration {
		public string Title { get; set; }
		public readonly Dictionary<string, AggregatedContainer> ColumnsMap = new Dictionary<string, AggregatedContainer>();
		public readonly List<ChildTableConfiguration> ChildrenMap = new List<ChildTableConfiguration>();
	}
}
