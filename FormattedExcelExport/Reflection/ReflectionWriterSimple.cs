using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using FormattedExcelExport.Style;
using FormattedExcelExport.TableWriters;


namespace FormattedExcelExport.Reflection {
	public static class ReflectionWriterSimple {
		public static MemoryStream Write<T>(IEnumerable<T> models, ITableWriterSimple tableWriter, CultureInfo cultureInfo) {
			IEnumerable<PropertyInfo> properties = typeof(T).GetProperties()
				.Where(x => x.PropertyType == typeof(string) || x.PropertyType == typeof(DateTime) || x.PropertyType == typeof(decimal) || x.PropertyType == typeof(int) || x.PropertyType == typeof(bool));

			var exportedProperties = new List<PropertyInfo>();
			var header = new List<string>();
			foreach (PropertyInfo propertyInfo in properties) {
				ExcelExportAttribute attribute = propertyInfo.GetCustomAttribute<ExcelExportAttribute>();
				if (attribute != null && !attribute.IsExportable) {
					continue;
				}
				header.Add(attribute != null ? attribute.Name : "");
				exportedProperties.Add(propertyInfo);
			}
			tableWriter.WriteHeader(header);

			foreach (T model in models) {
				var row = new List<string>();

				foreach (PropertyInfo propertyInfo in exportedProperties) {
					var propertyTypeName = propertyInfo.PropertyType.Name;

					switch (propertyTypeName) {
						case "String":
							row.Add(propertyInfo.GetValue(model).ToString());
							break;
						case "DateTime":
							row.Add(((DateTime)propertyInfo.GetValue(model)).ToString(cultureInfo.DateTimeFormat.LongDatePattern));
							break;
						case "Decimal":
							row.Add(string.Format(cultureInfo, "{0:C}", propertyInfo.GetValue(model)));
							break;
						case "Int32":
							row.Add(propertyInfo.GetValue(model).ToString());
							break;
						case "Boolean":
							row.Add(((bool)propertyInfo.GetValue(model)) ? "Да" : "Нет");
							break;
					}
				}
				tableWriter.WriteRow(row.ConvertAll(x => new KeyValuePair<string, TableWriterStyle>(x, null)));
			}

			tableWriter.AutosizeColumns();
			MemoryStream stream = tableWriter.GetStream();
			return stream;
		}
	}
}
