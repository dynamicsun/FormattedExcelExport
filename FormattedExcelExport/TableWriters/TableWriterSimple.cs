using System.Collections.Generic;
using System.IO;
using System.Linq;
using FormattedExcelExport.Configuaration;
using FormattedExcelExport.Infrastructure;
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
			List<string> parentNamesList = parentTableConfiguration.ColumnsMap.Keys.ToList();
			AggregatedContainer[] aggregatedContainers = parentTableConfiguration.ColumnsMap.Values.ToArray();
			List<ChildTableConfiguration> childTableConfigurations = parentTableConfiguration.ChildrenMap;

			int childTablesCount = childTableConfigurations.Count();
			var maximums = new List<int>();
			for (int i = 0; i < childTablesCount; i++) {
				maximums.Add(0);
			}

			foreach (TModel model in models) {
				int counter = 0;
				foreach (ChildTableConfiguration childTableConfiguration in childTableConfigurations) {
					IEnumerable<object> childTableRecords = childTableConfiguration.Getter(model);
					if (maximums[counter] < childTableRecords.Count())
						maximums[counter] = childTableRecords.Count();
					counter++;
				}
			}

			var childNamesList = new List<string>();
			int counter2 = 0;
			foreach (ChildTableConfiguration childTableConfiguration in childTableConfigurations) {
				int times = maximums[counter2];
				string[] keys = childTableConfiguration.ColumnsMap.Keys.ToArray();

				for (int i = 1; i <= times; i++) {
					foreach (string key in keys) {
						childNamesList.Add(key + i);
					}
				}
				counter2++;
			}

			parentNamesList.AddRange(childNamesList);
			writer.WriteHeader(parentNamesList.ToList());

			foreach (TModel model in models) {
                var cellsWithStyle = new List<KeyValuePair<dynamic, TableWriterStyle>>();

				foreach (AggregatedContainer aggregatedContainer in aggregatedContainers) {
					TableWriterStyle cellStyle = null;
					if (aggregatedContainer.ConditionFunc(model)) {
						cellStyle = aggregatedContainer.Style;
					}
					dynamic cell = aggregatedContainer.ValueFunc(model);
                    cellsWithStyle.Add(new KeyValuePair<dynamic, TableWriterStyle>(cell, cellStyle));
				}

				int counter3 = 0;
				foreach (ChildTableConfiguration childTableConfiguration in childTableConfigurations) {
					IEnumerable<object> children = childTableConfiguration.Getter(model);
					AggregatedContainer[] childAggregatedContainers = childTableConfiguration.ColumnsMap.Values.ToArray();

					foreach (object child in children) {
						foreach (AggregatedContainer childTableCellValueGetter in childAggregatedContainers) {
							TableWriterStyle cellStyle = null;
							if (childTableCellValueGetter.ConditionFunc(child)) {
								cellStyle = childTableCellValueGetter.Style;
							}
							dynamic cell = childTableCellValueGetter.ValueFunc(child);
                            cellsWithStyle.Add(new KeyValuePair<dynamic, TableWriterStyle>(cell, cellStyle));
						}
					}

					int difference = (maximums[counter3] - children.Count()) * childAggregatedContainers.Count();

					for (int i = 0; i < difference; i++) {
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
