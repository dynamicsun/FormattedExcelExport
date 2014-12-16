using System.Collections.Generic;
using System.IO;
using System.Linq;
using FormattedExcelExport.Configuaration;
using FormattedExcelExport.Style;


namespace FormattedExcelExport.TableWriters {
	public interface ITableWriterSimple {
		void WriteHeader(List<string> cells);
		void WriteRow(List<KeyValuePair<dynamic, TableWriterStyle>> cells);
		void AutosizeColumns();
		MemoryStream GetStream();
	}

	public static class TableWriterSimple {
		public static MemoryStream Write<TModel>(ITableWriterSimple writer, IEnumerable<TModel> models, TableConfiguration parentTableConfiguration) {
			var parentNamesList = parentTableConfiguration.ColumnsMap.Keys.ToList();
			var aggregatedContainers = parentTableConfiguration.ColumnsMap.Values.ToArray();
			var childTableConfigurations = parentTableConfiguration.ChildrenMap;

			var childTablesCount = childTableConfigurations.Count();
			var maximums = new List<int>();
			for (var i = 0; i < childTablesCount; i++) {
				maximums.Add(0);
			}

			foreach (var model in models) {
				var counter = 0;
				foreach (var childTableConfiguration in childTableConfigurations) {
					var childTableRecords = childTableConfiguration.Getter(model);
					if (maximums[counter] < childTableRecords.Count())
						maximums[counter] = childTableRecords.Count();
					counter++;
				}
			}

			var childNamesList = new List<string>();
			var counter2 = 0;
			foreach (var childTableConfiguration in childTableConfigurations) {
				var times = maximums[counter2];
				var keys = childTableConfiguration.ColumnsMap.Keys.ToArray();

				for (var i = 1; i <= times; i++) {
					foreach (var key in keys) {
						childNamesList.Add(key + i);
					}
				}
				counter2++;
			}

			parentNamesList.AddRange(childNamesList);
			writer.WriteHeader(parentNamesList.ToList());

			foreach (var model in models) {
                var cellsWithStyle = new List<KeyValuePair<dynamic, TableWriterStyle>>();

				foreach (var aggregatedContainer in aggregatedContainers) {
					TableWriterStyle cellStyle = null;
					if (aggregatedContainer.ConditionFunc(model)) {
						cellStyle = aggregatedContainer.Style;
					}
					var cell = aggregatedContainer.ValueFunc(model);
                    cellsWithStyle.Add(new KeyValuePair<dynamic, TableWriterStyle>(cell, cellStyle));
				}

				var counter3 = 0;
				foreach (var childTableConfiguration in childTableConfigurations) {
					var children = childTableConfiguration.Getter(model);
					var childAggregatedContainers = childTableConfiguration.ColumnsMap.Values.ToArray();

					foreach (var child in children) {
						foreach (var childTableCellValueGetter in childAggregatedContainers) {
							TableWriterStyle cellStyle = null;
							if (childTableCellValueGetter.ConditionFunc(child)) {
								cellStyle = childTableCellValueGetter.Style;
							}
							var cell = childTableCellValueGetter.ValueFunc(child);
                            cellsWithStyle.Add(new KeyValuePair<dynamic, TableWriterStyle>(cell, cellStyle));
						}
					}

					var difference = (maximums[counter3] - children.Count()) * childAggregatedContainers.Count();

					for (var i = 0; i < difference; i++) {
                        cellsWithStyle.Add(new KeyValuePair<dynamic, TableWriterStyle>(null, null));
					}
					counter3++;
				}

				writer.WriteRow(cellsWithStyle);
			}
			writer.AutosizeColumns();
			return writer.GetStream();
		}
	}
}
