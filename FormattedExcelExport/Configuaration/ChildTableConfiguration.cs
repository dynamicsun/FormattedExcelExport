using System;
using System.Collections.Generic;


namespace FormattedExcelExport.Configuaration {
	public class ChildTableConfiguration : TableConfiguration {
		public Func<object, IEnumerable<object>> Getter { get; set; }
	}
}