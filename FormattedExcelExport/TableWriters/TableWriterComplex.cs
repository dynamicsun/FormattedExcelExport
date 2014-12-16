using System.Collections.Generic;
using System.IO;
using System.Linq;
using FormattedExcelExport.Configuaration;


namespace FormattedExcelExport.TableWriters {
	public static class TableWriterComplex {
		public static MemoryStream Write<TModel>(ITableWriterComplex writer, IEnumerable<TModel> models, TableConfiguration parentTableConfiguration) {
			var headerNamesList = parentTableConfiguration.ColumnsMap.Keys.ToList();
			headerNamesList.Insert(0, parentTableConfiguration.Title);
			var aggregatedContainers = parentTableConfiguration.ColumnsMap.Values.ToArray();
			var childTableConfigurations = parentTableConfiguration.ChildrenMap;

			foreach (var model in models) {
				writer.WriteHeader(headerNamesList.ToArray());

			    var cellsWithStyle = TableWriterBase.AddCellStyles(aggregatedContainers, model);
				writer.WriteRow(cellsWithStyle.ToList(), true);

				foreach (var childTableConfiguration in childTableConfigurations) {
					var children = childTableConfiguration.Getter(model);
					var childAggregatedContainers = childTableConfiguration.ColumnsMap.Values.ToArray();
					var childHeaderNamesList = childTableConfiguration.ColumnsMap.Keys.ToList();
					childHeaderNamesList.Insert(0, childTableConfiguration.Title);
					writer.WriteChildHeader(childHeaderNamesList.ToArray());

					foreach (var child in children) {
					    var childCellsWithStyle = TableWriterBase.AddCellStyles(childAggregatedContainers, child);
						writer.WriteChildRow(childCellsWithStyle, true);
					}
				}
			}

			writer.AutosizeColumns();
			return writer.GetStream();
		}
	}
}
