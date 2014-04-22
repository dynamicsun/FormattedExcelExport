using System.Collections.Generic;
using System.IO;
using System.Linq;
using FormattedExcelExport.Configuaration;
using FormattedExcelExport.Infrastructure;
using FormattedExcelExport.Style;


namespace FormattedExcelExport.TableWriters {
	public static class TableWriterComplex {
		public static MemoryStream Write<TModel>(ITableWriterComplex writer, IEnumerable<TModel> models, TableConfiguration parentTableConfiguration) {
			List<string> headerNamesList = parentTableConfiguration.ColumnsMap.Keys.ToList();
			headerNamesList.Insert(0, parentTableConfiguration.Title);
			AggregatedContainer[] aggregatedContainers = parentTableConfiguration.ColumnsMap.Values.ToArray();
			List<ChildTableConfiguration> childTableConfigurations = parentTableConfiguration.ChildrenMap;

			foreach (TModel model in models) {
				writer.WriteHeader(headerNamesList.ToArray());

				var cellsWithStyle = new List<KeyValuePair<string, TableWriterStyle>>();
				foreach (AggregatedContainer aggregatedContainer in aggregatedContainers) {
					TableWriterStyle cellStyle = null;
					if (aggregatedContainer.ConditionFunc(model)) {
						cellStyle = aggregatedContainer.Style;
					}
					string cell = aggregatedContainer.ValueFunc(model);
					cellsWithStyle.Add(new KeyValuePair<string, TableWriterStyle>(cell, cellStyle));
				}
				writer.WriteRow(cellsWithStyle, true);

				foreach (ChildTableConfiguration childTableConfiguration in childTableConfigurations) {
					IEnumerable<object> children = childTableConfiguration.Getter(model);
					AggregatedContainer[] childAggregatedContainers = childTableConfiguration.ColumnsMap.Values.ToArray();
					List<string> childHeaderNamesList = childTableConfiguration.ColumnsMap.Keys.ToList();
					childHeaderNamesList.Insert(0, childTableConfiguration.Title);
					writer.WriteChildHeader(childHeaderNamesList.ToArray());

					foreach (object child in children) {
						var childCellsWithStyle = new List<KeyValuePair<string, TableWriterStyle>>();

						foreach (AggregatedContainer childTableCellValueGetter in childAggregatedContainers) {
							TableWriterStyle cellStyle = null;
							if (childTableCellValueGetter.ConditionFunc(child)) {
								cellStyle = childTableCellValueGetter.Style;
							}
							string cell = childTableCellValueGetter.ValueFunc(child);
							childCellsWithStyle.Add(new KeyValuePair<string, TableWriterStyle>(cell, cellStyle));
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
