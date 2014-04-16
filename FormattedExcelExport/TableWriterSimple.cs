using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;


namespace FormattedExcelExport {
	public interface ITableWriterSimple {
		void WriteHeader(params string[] cells);
		void WriteRow(params string[] cells);
		void AutosizeColumns();
		MemoryStream GetStream();
	}

	public static class TableWriterSimple {
		public static MemoryStream Write<TModel>(ITableWriterSimple writer, IEnumerable<TModel> models, TableConfiguration parentTableConfiguration, bool appendNumsToChildColumns = true) {
			List<string> headerNamesList = parentTableConfiguration.ColumnsMap.Keys.ToList();
			Func<object, string>[] parentTableCellValueGetters = parentTableConfiguration.ColumnsMap.Values.ToArray();
			ChildTableConfiguration childTableConfiguration = parentTableConfiguration.ChildrenMap.FirstOrDefault();

			if (childTableConfiguration == null) {
				writer.WriteHeader(headerNamesList.ToArray());

				foreach (TModel model in models) {
					var cells = new List<string>();

					foreach (Func<object, string> parentTableCellValueGetter in parentTableCellValueGetters) {
						string cell = parentTableCellValueGetter(model);
						cells.Add(cell);
					}
					writer.WriteRow(cells.ToArray());
					writer.WriteRow();
				}
			}
			else {
				Dictionary<string, Func<object, string>> childTableColumnsMap = childTableConfiguration.ColumnsMap;
				var childColumnNames = new List<string>();

				foreach (KeyValuePair<string, Func<object, string>> keyValuePair in childTableColumnsMap) {
					childColumnNames.Add(keyValuePair.Key);
				}

				Func<object, IEnumerable<object>> childTableCellValueGetters = childTableConfiguration.Getter;

				int maximumNestedChildrenCount = 0;
				foreach (TModel model in models) {
					IEnumerable<object> nestedChildren = childTableCellValueGetters(model);
					maximumNestedChildrenCount = maximumNestedChildrenCount > nestedChildren.Count() ? maximumNestedChildrenCount : nestedChildren.Count();
				}

				for (int i = 1; i <= maximumNestedChildrenCount; i++) {
					foreach (string childColumnName in childColumnNames) {
						if (appendNumsToChildColumns) {
							headerNamesList.Add(childColumnName + i);
						}
						else {
							headerNamesList.Add(childColumnName);
						}
					}
				}

				writer.WriteHeader(headerNamesList.ToArray());

				foreach (TModel model in models) {
					var cells = new List<string>();

					foreach (Func<object, string> keyValuePair in parentTableCellValueGetters) {
						string cell = keyValuePair(model);
						cells.Add(cell);
					}

					IEnumerable<object> childObjects = childTableCellValueGetters(model);
					foreach (object childObject in childObjects) {
						foreach (KeyValuePair<string, Func<object, string>> keyValuePair in childTableColumnsMap) {
							cells.Add(keyValuePair.Value(childObject));
						}
					}
					writer.WriteRow(cells.ToArray());
					writer.WriteRow();
				}
			}

			writer.AutosizeColumns();
			return writer.GetStream();
		}
	}
	public sealed class CsvTableWriterSimple : ITableWriterSimple {
		private readonly StringBuilder _stringBuilder = new StringBuilder();
		private readonly string _delimeter;
		public CsvTableWriterSimple(string delimeter = "\t") {
			_delimeter = delimeter;
		}

		public void WriteHeader(params string[] cells) {
			WriteRow(cells);
		}
		public void WriteRow(params string[] cells) {
			int cellsCount = cells.Length - 1;
			int i = 0;
			foreach (string cell in cells) {
				_stringBuilder.Append(cell);

				if (i < cellsCount)
					_stringBuilder.Append(_delimeter);
				i++;
			}
			_stringBuilder.AppendLine();
		}
		public void AutosizeColumns() { }
		public MemoryStream GetStream() {
			MemoryStream memoryStream = new MemoryStream();
			StreamWriter streamWriter = new StreamWriter(memoryStream, Encoding.UTF8);
			streamWriter.WriteLine(_stringBuilder.ToString());
			streamWriter.Flush();
			memoryStream.Position = 0;
			return memoryStream;
		}
	}
	public sealed class ExcelTableWriterSimple : ITableWriterSimple {
		private int _rowIndex;
		private readonly HSSFWorkbook _workbook;
		private readonly ISheet _workSheet;
		private readonly ITableWriterStyle _style;
		private ICellStyle _headerCellStyle;

		public ExcelTableWriterSimple(ITableWriterStyle style) {
			_style = style;
			_workbook = style.Workbook;
			_workSheet = _workbook.CreateSheet();
		}

		public void WriteHeader(params string[] cells) {
			IRow row = _workSheet.CreateRow(_rowIndex);
			row.Height = _style.HeaderHeight;
			int columnIndex = 0;
			foreach (var cell in cells) {
				ICell newCell = row.CreateCell(columnIndex);
				newCell.SetCellValue(cell);
				newCell.CellStyle = _style.HeaderCellStyle;
				columnIndex++;
			}
			_rowIndex++;
		}
		public void WriteRow(params string[] cells) {
			IRow newRow = _workSheet.CreateRow(_rowIndex);
			int columnIndex = 0;
			foreach (var cell in cells) {
				newRow.CreateCell(columnIndex).SetCellValue(cell);
				columnIndex++;
			}
			_rowIndex++;
		}

		public void AutosizeColumns() {
			var columnLengths = new List<int>();

			for (int columnNum = 0; columnNum < _workSheet.GetRow(0).LastCellNum; columnNum++) {
				int columnMaximumLength = 0;
				for (int rowNum = 0; rowNum < _workSheet.LastRowNum; rowNum++) {
					IRow currentRow = _workSheet.GetRow(rowNum);

					if (!currentRow.Cells.Any()) continue;
					ICell cell = currentRow.GetCell(columnNum);
					if (cell == null) continue;

					if (cell.StringCellValue.Length > columnMaximumLength)
						columnMaximumLength = cell.StringCellValue.Length;
				}
				columnLengths.Add(columnMaximumLength);
			}


			for (int i = 0; i < _workSheet.GetRow(0).LastCellNum; i++) {
				int width = columnLengths.ElementAt(i) * _style.FontFactor + _style.FontAbsoluteTerm;
				_workSheet.SetColumnWidth(i, width < _style.MaxColumnWidth ? width : _style.MaxColumnWidth);
			}
		}

		public MemoryStream GetStream() {
			MemoryStream memoryStream = new MemoryStream();
			_workbook.Write(memoryStream);
			memoryStream.Position = 0;
			return memoryStream;
		}
	}
}
