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
				.Where(x => x.PropertyType == typeof(string)
					|| x.PropertyType == typeof(DateTime) || x.PropertyType == typeof(DateTime?)
					|| x.PropertyType == typeof(decimal) || x.PropertyType == typeof(decimal?)
                    || x.PropertyType == typeof(float) || x.PropertyType == typeof(float?)
					|| x.PropertyType == typeof(int) || x.PropertyType == typeof(int?)
					|| x.PropertyType == typeof(bool) || x.PropertyType == typeof(bool?));

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

                List<KeyValuePair<dynamic, TableWriterStyle>> row = new List<KeyValuePair<dynamic, TableWriterStyle>>();
				ReflectionWriter.GetValue(cultureInfo, exportedProperties, row, model);
				tableWriter.WriteRow(row, true);
				row.Clear();

				for (int i = 0; i < enumerableProperties.Count(); i++) {
					PropertyInfo property = enumerableProperties.ElementAt(i);
					IList submodels = (IList)property.GetValue(model);

					Type propertyType = property.PropertyType;
					Type listType = propertyType.GetGenericArguments()[0];

					IEnumerable<PropertyInfo> props = listType.GetProperties()
					.Where(x => x.PropertyType == typeof(string)
						|| x.PropertyType == typeof(DateTime) || x.PropertyType == typeof(DateTime?)
						|| x.PropertyType == typeof(decimal) || x.PropertyType == typeof(decimal?)
                        || x.PropertyType == typeof(float) || x.PropertyType == typeof(float?)
						|| x.PropertyType == typeof(int) || x.PropertyType == typeof(int?)
						|| x.PropertyType == typeof(bool) || x.PropertyType == typeof(bool?));


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
						ReflectionWriter.GetValue(cultureInfo, props, row, submodel);
                        tableWriter.WriteChildRow(row, true);
						row.Clear();
					}
				}
			}

			tableWriter.AutosizeColumns();
			MemoryStream stream = tableWriter.GetStream();
			return stream;
		}
	}
}
