using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using FormattedExcelExport.Style;
using FormattedExcelExport.TableWriters;


namespace FormattedExcelExport.Reflection {
	public static class ReflectionWriterComplex {
		public static MemoryStream Write<T>(IEnumerable<T> models, ITableWriterComplex tableWriter, CultureInfo cultureInfo) {
			IEnumerable<PropertyInfo> nonEnumerableProperties = typeof(T).GetProperties()
				.Where(x => x.PropertyType == typeof(string) || x.PropertyType == typeof(DateTime) || x.PropertyType == typeof(decimal) || x.PropertyType == typeof(int) || x.PropertyType == typeof(bool));

			ExcelExportClassNameAttribute classAttribute = typeof(T).GetCustomAttribute<ExcelExportClassNameAttribute>();
			string className = classAttribute != null ? classAttribute.Name : "";

			var exportedProperties = new List<PropertyInfo>();
			var header = new List<string> { className };
			foreach (PropertyInfo propertyInfo in nonEnumerableProperties) {
				ExcelExportAttribute attribute = propertyInfo.GetCustomAttribute<ExcelExportAttribute>();
				if (attribute != null && !attribute.IsExportable) {
					continue;
				}
				header.Add(attribute != null ? attribute.PropertyName : "");
				exportedProperties.Add(propertyInfo);
			}

			IEnumerable<PropertyInfo> enumerableProperties = typeof(T).GetProperties().Where(x => x.PropertyType.IsGenericType && x.PropertyType.GetGenericTypeDefinition() == typeof(List<>));
			foreach (T model in models) {
				tableWriter.WriteHeader(header.ToArray());

				List<string> row = new List<string>();
				GetValue(cultureInfo, exportedProperties, row, model);
				tableWriter.WriteRow(row.ConvertAll(x => new KeyValuePair<string, TableWriterStyle>(x, null)), true);
				row.Clear();

				for (int i = 0; i < enumerableProperties.Count(); i++) {
					PropertyInfo property = enumerableProperties.ElementAt(i);
					IList submodels = (IList)property.GetValue(model);

					Type propertyType = property.PropertyType;
					Type listType = propertyType.GetGenericArguments()[0];

					IEnumerable<PropertyInfo> props = listType.GetProperties()
					.Where(x => x.PropertyType == typeof(string) || x.PropertyType == typeof(DateTime) || x.PropertyType == typeof(decimal) || x.PropertyType == typeof(int) || x.PropertyType == typeof(bool));


					ExcelExportClassNameAttribute nestedClassAttribute = listType.GetCustomAttribute<ExcelExportClassNameAttribute>();
					string nestedClassName = nestedClassAttribute != null ? nestedClassAttribute.Name : "";
					var childHeader = new List<string> { nestedClassName };
					foreach (PropertyInfo propertyInfo in props) {
						ExcelExportAttribute attribute = propertyInfo.GetCustomAttribute<ExcelExportAttribute>();
						if (attribute != null && !attribute.IsExportable) {
							continue;
						}
						childHeader.Add(attribute != null ? attribute.PropertyName : "");
					}
					tableWriter.WriteChildHeader(childHeader.ToArray());
					childHeader.Clear();

					foreach (var submodel in submodels) {
						GetValue(cultureInfo, props, row, submodel);
						tableWriter.WriteChildRow(row.ConvertAll(x => new KeyValuePair<string, TableWriterStyle>(x, null)), true);
						row.Clear();
					}
				}
			}

			tableWriter.AutosizeColumns();
			MemoryStream stream = tableWriter.GetStream();
			return stream;
		}

		private static void GetValue<T>(CultureInfo cultureInfo, IEnumerable<PropertyInfo> exportedProperties, List<string> row, T model) {
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
		}
	}
}
