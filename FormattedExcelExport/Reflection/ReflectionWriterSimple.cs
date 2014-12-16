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
	public static class ReflectionWriterSimple {
		public static MemoryStream Write<T>(IEnumerable<T> models, ITableWriterSimple tableWriter, CultureInfo cultureInfo) {
			var nonEnumerableProperties = typeof(T).GetProperties()
				.Where(x => x.PropertyType == typeof(string)
					|| x.PropertyType == typeof(DateTime) || x.PropertyType == typeof(DateTime?)
					|| x.PropertyType == typeof(decimal) || x.PropertyType == typeof(decimal?)
                    || x.PropertyType == typeof(float) || x.PropertyType == typeof(float?)
					|| x.PropertyType == typeof(int) || x.PropertyType == typeof(int?)
					|| x.PropertyType == typeof(bool) || x.PropertyType == typeof(bool?));

			var exportedProperties = new List<PropertyInfo>();
			var header = new List<string>();
			foreach (var propertyInfo in nonEnumerableProperties) {
				var attribute = propertyInfo.GetCustomAttribute<ExcelExportAttribute>();
				if (attribute == null || !attribute.IsExportable) {
					continue;
				}
				header.Add(attribute.PropertyName);
				exportedProperties.Add(propertyInfo);
			}

			var enumerableProperties = typeof(T).GetProperties().Where(x => x.PropertyType.IsGenericType && x.PropertyType.GetGenericTypeDefinition() == typeof(List<>));
			var maxims = new int[enumerableProperties.Count()];
			for (var i = 0; i < maxims.Count(); i++) {
				maxims[i] = 0;
			}
			foreach (var model in models) {
				for (var i = 0; i < enumerableProperties.Count(); i++) {
					var nestedModels = enumerableProperties.ElementAt(i).GetValue(model);
					if (maxims[i] < ((IList)nestedModels).Count) {
						maxims[i] = ((IList)nestedModels).Count;
					}
				}
			}

			for (var i = 0; i < enumerableProperties.Count(); i++) {
				var propertyType = enumerableProperties.ElementAt(i).PropertyType;
				var listType = propertyType.GetGenericArguments()[0];

				var props = listType.GetProperties()
					.Where(x => x.PropertyType == typeof(string)
						|| x.PropertyType == typeof(DateTime) || x.PropertyType == typeof(DateTime?)
						|| x.PropertyType == typeof(decimal) || x.PropertyType == typeof(decimal?)
                        || x.PropertyType == typeof(float) || x.PropertyType == typeof(float?)
						|| x.PropertyType == typeof(int) || x.PropertyType == typeof(int?)
						|| x.PropertyType == typeof(bool) || x.PropertyType == typeof(bool?));

				var counter = 1;
				for (var j = 0; j < maxims[i]; j++) {
					foreach (var propertyInfo in props) {
						var attribute = propertyInfo.GetCustomAttribute<ExcelExportAttribute>();
						if (attribute != null && !attribute.IsExportable) {
							continue;
						}
						header.Add(attribute != null ? attribute.PropertyName + counter : "");
					}
					counter++;
				}
			}

			tableWriter.WriteHeader(header);

			foreach (var model in models) {
                var row = new List<KeyValuePair<dynamic, TableWriterStyle>>();
				ReflectionWriter.GetValue(cultureInfo, exportedProperties, row, model);

				for (var i = 0; i < enumerableProperties.Count(); i++) {
					var property = enumerableProperties.ElementAt(i);
					var submodels = (IList)property.GetValue(model);
					
					var propertyType = property.PropertyType;
					var listType = propertyType.GetGenericArguments()[0];
					
					var props = listType.GetProperties()
					.Where(x => x.PropertyType == typeof(string)
						|| x.PropertyType == typeof(DateTime) || x.PropertyType == typeof(DateTime?)
						|| x.PropertyType == typeof(decimal) || x.PropertyType == typeof(decimal?)
                        || x.PropertyType == typeof(float) || x.PropertyType == typeof(float?)
						|| x.PropertyType == typeof(int) || x.PropertyType == typeof(int?)
						|| x.PropertyType == typeof(bool) || x.PropertyType == typeof(bool?));

					foreach (var submodel in submodels) {
						ReflectionWriter.GetValue(cultureInfo, props, row, submodel);
					}
					for (var j = 0; j < (maxims[i] - submodels.Count) * props.Count(); j++) {
                        row.Add(new KeyValuePair<dynamic, TableWriterStyle>(string.Empty, null));
					}
				}
                tableWriter.WriteRow(row);
			}

			tableWriter.AutosizeColumns();
			var stream = tableWriter.GetStream();
			return stream;
		}
	}
}
