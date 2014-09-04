using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using FormattedExcelExport.Style;
using FormattedExcelExport.TableWriters;


namespace FormattedExcelExport.Reflection {
	public static class ReflectionWriterSimple {
		public static MemoryStream Write<T>(IEnumerable<T> models, ITableWriterSimple tableWriter, CultureInfo cultureInfo) {
			IEnumerable<PropertyInfo> nonEnumerableProperties = typeof(T).GetProperties()
				.Where(x => x.PropertyType == typeof(string)
					|| x.PropertyType == typeof(DateTime) || x.PropertyType == typeof(DateTime?)
					|| x.PropertyType == typeof(decimal) || x.PropertyType == typeof(decimal?)
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
                var row = new List<KeyValuePair<string, TableWriterStyle>>();
				GetValue(cultureInfo, exportedProperties, row, model);

				for (int i = 0; i < enumerableProperties.Count(); i++) {
					var property = enumerableProperties.ElementAt(i);
					IList submodels = (IList)property.GetValue(model);
					
					var propertyType = property.PropertyType;
					Type listType = propertyType.GetGenericArguments()[0];
					
					IEnumerable<PropertyInfo> props = listType.GetProperties()
					.Where(x => x.PropertyType == typeof(string)
						|| x.PropertyType == typeof(DateTime) || x.PropertyType == typeof(DateTime?)
						|| x.PropertyType == typeof(decimal) || x.PropertyType == typeof(decimal?)
						|| x.PropertyType == typeof(int) || x.PropertyType == typeof(int?)
						|| x.PropertyType == typeof(bool) || x.PropertyType == typeof(bool?));

					foreach (var submodel in submodels) {
						GetValue(cultureInfo, props, row, submodel);
					}
					for (int j = 0; j < (maxims[i] - submodels.Count) * props.Count(); j++) {
                        row.Add(new KeyValuePair<string, TableWriterStyle> ("", null));
					}
				}
                tableWriter.WriteRow(row);
				//tableWriter.WriteRow(row.ConvertAll(x => new KeyValuePair<string, TableWriterStyle>(x, null)));
			}

			tableWriter.AutosizeColumns();
			MemoryStream stream = tableWriter.GetStream();
			return stream;
		}
        private static void GetValue<T>(CultureInfo cultureInfo, IEnumerable<PropertyInfo> exportedProperties, List<KeyValuePair<string, TableWriterStyle>> row, T model) {
            TableWriterStyle style = new TableWriterStyle();
            style.RegularCell.BackgroundColor = new AdHocCellStyle.Color(255, 0, 0);
            style.RegularChildCell.BackgroundColor = new AdHocCellStyle.Color(255, 0, 0);
			foreach (PropertyInfo propertyInfo in exportedProperties) {
				var propertyTypeName = propertyInfo.PropertyType.Name;
			    if (propertyInfo.GetValue(model) == null) {
			        row.Add(new KeyValuePair<string, TableWriterStyle>(String.Empty, style));
			    }
				switch (propertyTypeName) {
					case "String":
				        if (propertyInfo.GetCustomAttribute<ExcelExportAttribute>().ConditionType != null) {
                            row.Add(new KeyValuePair<string, TableWriterStyle>(propertyInfo.GetValue(model).ToString(), style));
				        }
				        else {
                            row.Add(new KeyValuePair<string, TableWriterStyle>(
                                propertyInfo.GetValue(model) == null 
                                    ? string.Empty 
                                    : propertyInfo.GetValue(model).ToString(), 
                                null));
				        }
						//row.Add(propertyInfo.GetValue(model).ToString());
						break;
					case "DateTime":
                        Thread.CurrentThread.CurrentCulture = cultureInfo;
                        if (propertyInfo.GetCustomAttribute<ExcelExportAttribute>().ConditionType != null) {
                            row.Add(new KeyValuePair<string, TableWriterStyle>(((DateTime)propertyInfo.GetValue(model)).ToString(cultureInfo.DateTimeFormat.LongDatePattern), style));
                        } else {
                            row.Add(new KeyValuePair<string, TableWriterStyle>(((DateTime)propertyInfo.GetValue(model)).ToString(cultureInfo.DateTimeFormat.LongDatePattern), null));
                        }
						//row.Add(((DateTime) propertyInfo.GetValue(model)).ToString(cultureInfo.DateTimeFormat.LongDatePattern));
						break;
					case "Decimal":
				        Type conditionType = propertyInfo.GetCustomAttribute<ExcelExportAttribute>().ConditionType;
                        if (conditionType != null) {
                            row.Add(new KeyValuePair<string, TableWriterStyle>(string.Format(cultureInfo, "{0:C}", propertyInfo.GetValue(model)), style));
                        } else {
                            row.Add(new KeyValuePair<string, TableWriterStyle>(string.Format(cultureInfo, "{0:C}", propertyInfo.GetValue(model)), null));
                        }
						//row.Add(string.Format(cultureInfo, "{0:C}", propertyInfo.GetValue(model)));
						break;
					case "Int32":
                        if (propertyInfo.GetCustomAttribute<ExcelExportAttribute>().ConditionType != null) {
                            row.Add(new KeyValuePair<string, TableWriterStyle>(propertyInfo.GetValue(model).ToString(), style));
                        } else {
                            row.Add(new KeyValuePair<string, TableWriterStyle>(propertyInfo.GetValue(model).ToString(), null));
                        }
						//row.Add(propertyInfo.GetValue(model).ToString());
						break;
					case "Boolean":
                        if (propertyInfo.GetCustomAttribute<ExcelExportAttribute>().ConditionType != null) {
                            row.Add(new KeyValuePair<string, TableWriterStyle>(((bool)propertyInfo.GetValue(model)) ? "Да" : "Нет", style));
                        } else {
                            row.Add(new KeyValuePair<string, TableWriterStyle>(((bool)propertyInfo.GetValue(model)) ? "Да" : "Нет", null));
                        }
						//row.Add(((bool) propertyInfo.GetValue(model)) ? "Да" : "Нет");
						break;
					case "Nullable`1":
				        if (propertyInfo.PropertyType.FullName.Contains("DateTime")) {
				            if (propertyInfo.GetCustomAttribute<ExcelExportAttribute>().ConditionType != null) {
				                row.Add(new KeyValuePair<string, TableWriterStyle>(((DateTime) propertyInfo.GetValue(model)).ToString(cultureInfo.DateTimeFormat.LongDatePattern), style));
				            } else {
				                row.Add(new KeyValuePair<string, TableWriterStyle>(((DateTime) propertyInfo.GetValue(model)).ToString(cultureInfo.DateTimeFormat.LongDatePattern), null));
				            }
				        }
				        if (propertyInfo.PropertyType.FullName.Contains("Decimal")) {
				            if (propertyInfo.GetCustomAttribute<ExcelExportAttribute>().ConditionType != null) {
				                row.Add(new KeyValuePair<string, TableWriterStyle>(string.Format(cultureInfo, "{0}", propertyInfo.GetValue(model)), style));
				            } else {
                                row.Add(new KeyValuePair<string, TableWriterStyle>(string.Format(cultureInfo, "{0}", propertyInfo.GetValue(model)), null));
				            }
				        }
				        if (propertyInfo.PropertyType.FullName.Contains("Int32")) {
				            if (propertyInfo.GetCustomAttribute<ExcelExportAttribute>().ConditionType != null) {
				                row.Add(new KeyValuePair<string, TableWriterStyle>(propertyInfo.GetValue(model).ToString(), style));
				            } else {
                                row.Add(new KeyValuePair<string, TableWriterStyle>(propertyInfo.GetValue(model).ToString(), null));
				            }
				        }
				        if (propertyInfo.PropertyType.FullName.Contains("Boolean")) {
				            if (propertyInfo.GetCustomAttribute<ExcelExportAttribute>().ConditionType != null) {
				                row.Add(new KeyValuePair<string, TableWriterStyle>(((bool) propertyInfo.GetValue(model)) ? "Да" : "Нет", style));
				            } else {
                                row.Add(new KeyValuePair<string, TableWriterStyle>(((bool)propertyInfo.GetValue(model)) ? "Да" : "Нет", null));
				            }
				        }
				        break;
				}
			}
		}
	}
}
