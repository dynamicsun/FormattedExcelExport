using System;


namespace FormattedExcelExport.Configuaration {
	[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
	public sealed class ExcelExportAttribute : Attribute {
		public string Name { get; set; }
		public bool? IsExportable { get; set; }
	}
}

