using System;


namespace FormattedExcelExport.Reflection {
	[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
	public sealed class ExcelExportAttribute : Attribute {
		public ExcelExportAttribute() {
			PropertyName = "";
			IsExportable = true;
		}
		public ExcelExportAttribute(string propertyName = "", bool isExportable = true) {
			PropertyName = propertyName;
			IsExportable = isExportable;
		}
		public string PropertyName { get; set; }
		public bool IsExportable { get; set; }
	}
	[AttributeUsage(AttributeTargets.Class)]
	public sealed class ExcelExportClassNameAttribute : Attribute {
		public ExcelExportClassNameAttribute() {
		}
		public ExcelExportClassNameAttribute(string name = "") {
			Name = name;
		}
		public string Name { get; set; }
	}

}

