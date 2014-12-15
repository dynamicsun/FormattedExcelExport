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
			IEnumerable<PropertyInfo> nonEnumerableProperties = typeof(T).GetProperties()
				.Where(x => x.PropertyType == typeof(string)
					|| x.PropertyType == typeof(DateTime) || x.PropertyType == typeof(DateTime?)
					|| x.PropertyType == typeof(decimal) || x.PropertyType == typeof(decimal?)
                    || x.PropertyType == typeof(float) || x.PropertyType == typeof(float?)
					|| x.PropertyType == typeof(int) || x.PropertyType == typeof(int?)
					|| x.PropertyType == typeof(bool) || x.PropertyType == typeof(bool?));

			var exportedProperties = new List<PropertyInfo>();
			var header = new List<string>();
			foreach (PropertyInfo propertyInfo in nonEnumerableProperties) {
				ExcelExportAttribute attribute = propertyInfo.GetCustomAttribute<ExcelExportAttribute>();
				if (attribute == null || !attribute.IsExportable) {
					continue;
				}
				header.Add(attribute.PropertyName);
				exportedProperties.Add(propertyInfo);
			}

			IEnumerable<PropertyInfo> enumerableProperties = typeof(T).GetProperties().Where(x => x.PropertyType.IsGenericType && x.PropertyType.GetGenericTypeDefinition() == typeof(List<>));
			int[] maxims = new int[enumerableProperties.Count()];
			for (int i = 0; i < maxims.Count(); i++) {
				maxims[i] = 0;
			}
			foreach (T model in models) {
				for (int i = 0; i < enumerableProperties.Count(); i++) {
					object nestedModels = enumerableProperties.ElementAt(i).GetValue(model);
					if (maxims[i] < ((IList)nestedModels).Count) {
						maxims[i] = ((IList)nestedModels).Count;
					}
				}
			}

			for (int i = 0; i < enumerableProperties.Count(); i++) {
				Type propertyType = enumerableProperties.ElementAt(i).PropertyType;
				Type listType = propertyType.GetGenericArguments()[0];

				IEnumerable<PropertyInfo> props = listType.GetProperties()
					.Where(x => x.PropertyType == typeof(string)
						|| x.PropertyType == typeof(DateTime) || x.PropertyType == typeof(DateTime?)
						|| x.PropertyType == typeof(decimal) || x.PropertyType == typeof(decimal?)
                        || x.PropertyType == typeof(float) || x.PropertyType == typeof(float?)
						|| x.PropertyType == typeof(int) || x.PropertyType == typeof(int?)
						|| x.PropertyType == typeof(bool) || x.PropertyType == typeof(bool?));

				int counter = 1;
				for (int j = 0; j < maxims[i]; j++) {
					foreach (PropertyInfo propertyInfo in props) {
						ExcelExportAttribute attribute = propertyInfo.GetCustomAttribute<ExcelExportAttribute>();
						if (attribute != null && !attribute.IsExportable) {
							continue;
						}
						header.Add(attribute != null ? attribute.PropertyName + counter : "");
					}
					counter++;
				}
			}

			tableWriter.WriteHeader(header);

			foreach (T model in models) {
                var row = new List<KeyValuePair<dynamic, TableWriterStyle>>();
				ReflectionWriter.GetValue(cultureInfo, exportedProperties, row, model);

				for (int i = 0; i < enumerableProperties.Count(); i++) {
					var property = enumerableProperties.ElementAt(i);
					IList submodels = (IList)property.GetValue(model);
					
					var propertyType = property.PropertyType;
					Type listType = propertyType.GetGenericArguments()[0];
					
					IEnumerable<PropertyInfo> props = listType.GetProperties()
					.Where(x => x.PropertyType == typeof(string)
						|| x.PropertyType == typeof(DateTime) || x.PropertyType == typeof(DateTime?)
						|| x.PropertyType == typeof(decimal) || x.PropertyType == typeof(decimal?)
                        || x.PropertyType == typeof(float) || x.PropertyType == typeof(float?)
						|| x.PropertyType == typeof(int) || x.PropertyType == typeof(int?)
						|| x.PropertyType == typeof(bool) || x.PropertyType == typeof(bool?));

					foreach (var submodel in submodels) {
						ReflectionWriter.GetValue(cultureInfo, props, row, submodel);
					}
					for (int j = 0; j < (maxims[i] - submodels.Count) * props.Count(); j++) {
                        row.Add(new KeyValuePair<dynamic, TableWriterStyle>(string.Empty, null));
					}
				}
                tableWriter.WriteRow(row);
			}

			tableWriter.AutosizeColumns();
			MemoryStream stream = tableWriter.GetStream();
			return stream;
		}
	}
}
