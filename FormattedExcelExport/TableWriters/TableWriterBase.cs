using System.Collections.Generic;
using FormattedExcelExport.Infrastructure;
using FormattedExcelExport.Style;

namespace FormattedExcelExport.TableWriters {
    public static class TableWriterBase {
        public static IEnumerable<KeyValuePair<object, TableWriterStyle>> AddCellStyles<TModel>(AggregatedContainer[] aggregatedContainers, TModel model) {
            var cellsWithStyle = new List<KeyValuePair<object, TableWriterStyle>>();
            foreach (var aggregatedContainer in aggregatedContainers) {
                TableWriterStyle cellStyle = null;
                if (aggregatedContainer.ConditionFunc(model)) {
                    cellStyle = aggregatedContainer.Style;
                }
                var cell = aggregatedContainer.ValueFunc(model);
                cellsWithStyle.Add(new KeyValuePair<object, TableWriterStyle>(cell, cellStyle));
            }
            return cellsWithStyle;
        }
    }
}
