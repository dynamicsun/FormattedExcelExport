using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using FormattedExcelExport.Style;

namespace FormattedExcelExport.Reflection {
    public static class ReflectionWriter {
        public static IEnumerable<PropertyInfo> ReflectionGetProperties(Type type) {
            return type.GetProperties()
                .Where(x => x.PropertyType == typeof(string)
                            || x.PropertyType == typeof(DateTime) || x.PropertyType == typeof(DateTime?)
                            || x.PropertyType == typeof(decimal) || x.PropertyType == typeof(decimal?)
                            || x.PropertyType == typeof(float) || x.PropertyType == typeof(float?)
                            || x.PropertyType == typeof(int) || x.PropertyType == typeof(int?)
                            || x.PropertyType == typeof(bool) || x.PropertyType == typeof(bool?));
        }
        public static void GetValue<T>(CultureInfo cultureInfo, IEnumerable<PropertyInfo> exportedProperties, List<KeyValuePair<dynamic, TableWriterStyle>> row, T model) {
            var style = new TableWriterStyle();
            style.RegularCell.BackgroundColor = new AdHocCellStyle.Color(255, 0, 0);
            style.RegularChildCell.BackgroundColor = new AdHocCellStyle.Color(255, 0, 0);
            foreach (var propertyInfo in exportedProperties) {
				if (propertyInfo.GetCustomAttribute<ExcelExportAttribute>() == null) continue;

                if (propertyInfo.PropertyType.FullName.Contains("DateTime")) {
                    if (propertyInfo.GetCustomAttribute<ExcelExportAttribute>().ConditionType != null) {
                        row.Add(new KeyValuePair<dynamic, TableWriterStyle>((DateTime?)propertyInfo.GetValue(model), style));
                    } else {
                        row.Add(new KeyValuePair<dynamic, TableWriterStyle>((DateTime?)propertyInfo.GetValue(model), null));
                    }
                }
                if (propertyInfo.PropertyType.FullName.Contains("Decimal") || propertyInfo.PropertyType.FullName.Contains("Single") || propertyInfo.PropertyType.FullName.Contains("Int32")) {
                    if (propertyInfo.GetCustomAttribute<ExcelExportAttribute>().ConditionType != null) {
                        row.Add(new KeyValuePair<dynamic, TableWriterStyle>(propertyInfo.GetValue(model) != null ? (double?)Convert.ToDouble(propertyInfo.GetValue(model)) : null, style));
                    } else {
                        row.Add(new KeyValuePair<dynamic, TableWriterStyle>(propertyInfo.GetValue(model) != null ? (double?)Convert.ToDouble(propertyInfo.GetValue(model)) : null, null));
                    }
                }
                if (propertyInfo.PropertyType.FullName.Contains("Boolean")) {
                    var value = (bool?)propertyInfo.GetValue(model);
                    if (propertyInfo.GetCustomAttribute<ExcelExportAttribute>().ConditionType != null) {
                        row.Add(new KeyValuePair<dynamic, TableWriterStyle>(value.HasValue ? (value.Value ? "Да" : "Нет") : string.Empty, style));
                    } else {
                        row.Add(new KeyValuePair<dynamic, TableWriterStyle>(value.HasValue ? (value.Value ? "Да" : "Нет") : string.Empty, null));
                    }
                }
                if (propertyInfo.PropertyType.FullName.Contains("String")) {
                    if (propertyInfo.GetCustomAttribute<ExcelExportAttribute>().ConditionType != null) {
                        row.Add(new KeyValuePair<dynamic, TableWriterStyle>((string)propertyInfo.GetValue(model), style));
                    } else {
                        row.Add(new KeyValuePair<dynamic, TableWriterStyle>((string)propertyInfo.GetValue(model), null));
                    }
                }
            }
        }
    }
}
