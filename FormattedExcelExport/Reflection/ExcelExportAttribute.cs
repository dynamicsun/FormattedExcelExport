using System;


namespace FormattedExcelExport.Reflection {
	[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
	public sealed class ExcelExportAttribute : Attribute {
		public ExcelExportAttribute() {
			Name = "";
			IsExportable = true;
		}
		public ExcelExportAttribute(string name = "", bool isExportable = true) {
			Name = name;
			IsExportable = isExportable;
		}
		public string Name { get; set; }
		public bool IsExportable { get; set; }
	}
}

