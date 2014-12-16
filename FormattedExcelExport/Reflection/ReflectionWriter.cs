using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Threading;
using FormattedExcelExport.Style;

namespace FormattedExcelExport.Reflection {
    public static class ReflectionWriter {
        public static void GetValue<T>(CultureInfo cultureInfo, IEnumerable<PropertyInfo> exportedProperties, List<KeyValuePair<dynamic, TableWriterStyle>> row, T model) {
            var style = new TableWriterStyle();
            style.RegularCell.BackgroundColor = new AdHocCellStyle.Color(255, 0, 0);
            style.RegularChildCell.BackgroundColor = new AdHocCellStyle.Color(255, 0, 0);
            foreach (var propertyInfo in exportedProperties) {
                var propertyTypeName = propertyInfo.PropertyType.Name;
                bool? value;
                switch (propertyTypeName) {
                    case "String":
                        if (propertyInfo.GetCustomAttribute<ExcelExportAttribute>().ConditionType != null) {
                            row.Add(new KeyValuePair<dynamic, TableWriterStyle>((string)propertyInfo.GetValue(model), style));
                        } else {
                            row.Add(new KeyValuePair<dynamic, TableWriterStyle>((string)propertyInfo.GetValue(model), null));
                        }
                        break;
                    case "DateTime":
                        if (propertyInfo.GetCustomAttribute<ExcelExportAttribute>().ConditionType != null) {
                            row.Add(new KeyValuePair<dynamic, TableWriterStyle>((DateTime?)propertyInfo.GetValue(model), style));
                        } else {
                            row.Add(new KeyValuePair<dynamic, TableWriterStyle>((DateTime?)propertyInfo.GetValue(model), null));
                        }
                        break;
                    case "Decimal":
                        var conditionType = propertyInfo.GetCustomAttribute<ExcelExportAttribute>().ConditionType;
                        if (conditionType != null) {
                            row.Add(new KeyValuePair<dynamic, TableWriterStyle>(propertyInfo.GetValue(model) != null ? (double?)Convert.ToDouble(propertyInfo.GetValue(model)) : null, style));
                        } else {
                            row.Add(new KeyValuePair<dynamic, TableWriterStyle>(propertyInfo.GetValue(model) != null ? (double?)Convert.ToDouble(propertyInfo.GetValue(model)) : null, null));
                        }
                        break;
                    case "Single":
                        if (propertyInfo.GetCustomAttribute<ExcelExportAttribute>().ConditionType != null) {
                            row.Add(new KeyValuePair<dynamic, TableWriterStyle>(propertyInfo.GetValue(model) != null ? (double?)Convert.ToDouble(propertyInfo.GetValue(model)) : null, style));
                        } else {
                            row.Add(new KeyValuePair<dynamic, TableWriterStyle>(propertyInfo.GetValue(model) != null ? (double?)Convert.ToDouble(propertyInfo.GetValue(model)) : null, null));
                        }
                        break;
                    case "Int32":
                        if (propertyInfo.GetCustomAttribute<ExcelExportAttribute>().ConditionType != null) {
                            row.Add(new KeyValuePair<dynamic, TableWriterStyle>((int?)propertyInfo.GetValue(model), style));
                        } else {
                            row.Add(new KeyValuePair<dynamic, TableWriterStyle>((int?)propertyInfo.GetValue(model), null));
                        }
                        break;
                    case "Boolean":
                        value = (bool?)propertyInfo.GetValue(model);
                        if (propertyInfo.GetCustomAttribute<ExcelExportAttribute>().ConditionType != null) {
                            row.Add(new KeyValuePair<dynamic, TableWriterStyle>(value.HasValue ? (value.Value ? "Да" : "Нет") : string.Empty, style));
                        } else {
                            row.Add(new KeyValuePair<dynamic, TableWriterStyle>(value.HasValue ? (value.Value ? "Да" : "Нет") : string.Empty, null));
                        }
                        break;
                    case "Nullable`1":
                        if (propertyInfo.PropertyType.FullName.Contains("DateTime")) {
                            Thread.CurrentThread.CurrentCulture = cultureInfo;
                            if (propertyInfo.GetCustomAttribute<ExcelExportAttribute>().ConditionType != null) {
                                row.Add(new KeyValuePair<dynamic, TableWriterStyle>((DateTime?)propertyInfo.GetValue(model), style));
                            } else {
                                row.Add(new KeyValuePair<dynamic, TableWriterStyle>((DateTime?)propertyInfo.GetValue(model), null));
                            }
                        }
                        if (propertyInfo.PropertyType.FullName.Contains("Decimal")) {
                            if (propertyInfo.GetCustomAttribute<ExcelExportAttribute>().ConditionType != null) {
                                row.Add(new KeyValuePair<dynamic, TableWriterStyle>(propertyInfo.GetValue(model) != null ? (double?)Convert.ToDouble(propertyInfo.GetValue(model)) : null, style));
                            } else {
                                row.Add(new KeyValuePair<dynamic, TableWriterStyle>(propertyInfo.GetValue(model) != null ? (double?)Convert.ToDouble(propertyInfo.GetValue(model)) : null, null));
                            }
                        }
                        if (propertyInfo.PropertyType.FullName.Contains("Single")) {
                            if (propertyInfo.GetCustomAttribute<ExcelExportAttribute>().ConditionType != null) {
                                row.Add(new KeyValuePair<dynamic, TableWriterStyle>(propertyInfo.GetValue(model) != null ? (double?)Convert.ToDouble(propertyInfo.GetValue(model)) : null, style));
                            } else {
                                row.Add(new KeyValuePair<dynamic, TableWriterStyle>(propertyInfo.GetValue(model) != null ? (double?)Convert.ToDouble(propertyInfo.GetValue(model)) : null, null));
                            }
                        }
                        if (propertyInfo.PropertyType.FullName.Contains("Int32")) {
                            if (propertyInfo.GetCustomAttribute<ExcelExportAttribute>().ConditionType != null) {
                                row.Add(new KeyValuePair<dynamic, TableWriterStyle>((int?)propertyInfo.GetValue(model), style));
                            } else {
                                row.Add(new KeyValuePair<dynamic, TableWriterStyle>((int?)propertyInfo.GetValue(model), null));
                            }
                        }
                        if (propertyInfo.PropertyType.FullName.Contains("Boolean")) {
                            value = (bool?)propertyInfo.GetValue(model);
                            if (propertyInfo.GetCustomAttribute<ExcelExportAttribute>().ConditionType != null) {
                                row.Add(new KeyValuePair<dynamic, TableWriterStyle>(value.HasValue ? (value.Value ? "Да" : "Нет") : string.Empty, style));
                            } else {
                                row.Add(new KeyValuePair<dynamic, TableWriterStyle>(value.HasValue ? (value.Value ? "Да" : "Нет") : string.Empty, null));
                            }
                        }
                        break;
                }
            }
        }
    }
}
