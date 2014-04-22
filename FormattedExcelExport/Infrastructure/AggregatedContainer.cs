using System;
using FormattedExcelExport.Style;


namespace FormattedExcelExport.Infrastructure {
	public class AggregatedContainer {
		public AggregatedContainer(Func<object, string> valueFunc, Func<object, bool> conditionFunc, TableWriterStyle style) {
			ValueFunc = valueFunc;
			ConditionFunc = conditionFunc;
			Style = style;
		}
		public Func<object, string> ValueFunc { get; set; }
		public TableWriterStyle Style { get; set; }
		public Func<object, bool> ConditionFunc { get; set; }
	}
}