using System.Collections.Generic;
using System.IO;
using System.Linq;
using FormattedExcelExport.Configuaration;
using FormattedExcelExport.Infrastructure;
using FormattedExcelExport.Style;


namespace FormattedExcelExport.TableWriters {
	public static class TableWriterComplex {
		public static MemoryStream Write<TModel>(ITableWriterComplex writer, IEnumerable<TModel> models, TableConfiguration parentTableConfiguration) {
			var headerNamesList = parentTableConfiguration.ColumnsMap.Keys.ToList();
			headerNamesList.Insert(0, parentTableConfiguration.Title);
			var aggregatedContainers = parentTableConfiguration.ColumnsMap.Values.ToArray();
			var childTableConfigurations = parentTableConfiguration.ChildrenMap;

			foreach (var model in models) {
				writer.WriteHeader(headerNamesList.ToArray());

                var cellsWithStyle = new List<KeyValuePair<dynamic, TableWriterStyle>>();
				foreach (var aggregatedContainer in aggregatedContainers) {
					TableWriterStyle cellStyle = null;
					if (aggregatedContainer.ConditionFunc(model)) {
						cellStyle = aggregatedContainer.Style;
					}
					var cell = aggregatedContainer.ValueFunc(model);
                    cellsWithStyle.Add(new KeyValuePair<dynamic, TableWriterStyle>(cell, cellStyle));
				}
				writer.WriteRow(cellsWithStyle, true);

				foreach (var childTableConfiguration in childTableConfigurations) {
					var children = childTableConfiguration.Getter(model);
					var childAggregatedContainers = childTableConfiguration.ColumnsMap.Values.ToArray();
					var childHeaderNamesList = childTableConfiguration.ColumnsMap.Keys.ToList();
					childHeaderNamesList.Insert(0, childTableConfiguration.Title);
					writer.WriteChildHeader(childHeaderNamesList.ToArray());

					foreach (var child in children) {
                        var childCellsWithStyle = new List<KeyValuePair<dynamic, TableWriterStyle>>();

						foreach (var childTableCellValueGetter in childAggregatedContainers) {
							TableWriterStyle cellStyle = null;
							if (childTableCellValueGetter.ConditionFunc(child)) {
								cellStyle = childTableCellValueGetter.Style;
							}
							var cell = childTableCellValueGetter.ValueFunc(child);
                            childCellsWithStyle.Add(new KeyValuePair<dynamic, TableWriterStyle>(cell, cellStyle));
						}
						writer.WriteChildRow(childCellsWithStyle, true);
					}
				}
			}

			writer.AutosizeColumns();
			return writer.GetStream();
		}
	}
}
