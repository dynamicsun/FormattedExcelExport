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
            var nonEnumerableProperties = ReflectionWriter.ReflectionGetProperties(typeof(T));

			var classAttribute = typeof(T).GetCustomAttribute<ExcelExportClassNameAttribute>();
			var className = classAttribute != null ? classAttribute.Name : "";

			var exportedProperties = new List<PropertyInfo>();
			var header = new List<string> { className };
			foreach (var propertyInfo in nonEnumerableProperties) {
				var attribute = propertyInfo.GetCustomAttribute<ExcelExportAttribute>();
				if (attribute != null && !attribute.IsExportable) {
					continue;
				}
				header.Add(attribute != null ? attribute.PropertyName : "");
				exportedProperties.Add(propertyInfo);
			}

			var enumerableProperties = typeof(T).GetProperties().Where(x => x.PropertyType.IsGenericType && x.PropertyType.GetGenericTypeDefinition() == typeof(List<>));
			foreach (var model in models) {
				tableWriter.WriteHeader(header.ToArray());

                var row = new List<KeyValuePair<dynamic, TableWriterStyle>>();
				ReflectionWriter.GetValue(cultureInfo, exportedProperties, row, model);
				tableWriter.WriteRow(row, true);
				row.Clear();

				for (var i = 0; i < enumerableProperties.Count(); i++) {
					var property = enumerableProperties.ElementAt(i);
					var submodels = (IList)property.GetValue(model);

					var propertyType = property.PropertyType;
					var listType = propertyType.GetGenericArguments()[0];
                    var props = ReflectionWriter.ReflectionGetProperties(listType);
					var nestedClassAttribute = listType.GetCustomAttribute<ExcelExportClassNameAttribute>();
					var nestedClassName = nestedClassAttribute != null ? nestedClassAttribute.Name : "";
					var childHeader = new List<string> { nestedClassName };
					foreach (var propertyInfo in props) {
						var attribute = propertyInfo.GetCustomAttribute<ExcelExportAttribute>();
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
			var stream = tableWriter.GetStream();
			return stream;
		}
	}
}
