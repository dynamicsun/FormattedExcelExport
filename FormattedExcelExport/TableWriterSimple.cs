using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;


namespace FormattedExcelExport {
	public interface ITableWriterSimple {
		void WriteHeader(params string[] cells);
		void WriteRow(List<KeyValuePair<string, TableWriterStyle>> cells);
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
			writer.WriteHeader(parentNamesList.ToArray());

			foreach (TModel model in models) {
				var cellsWithStyle = new List<KeyValuePair<string, TableWriterStyle>>();

				foreach (AggregatedContainer aggregatedContainer in aggregatedContainers) {
					TableWriterStyle cellStyle = null;
					if (aggregatedContainer.ConditionFunc(model)) {
						cellStyle = aggregatedContainer.Style;
					}
					string cell = aggregatedContainer.ValueFunc(model);
					cellsWithStyle.Add(new KeyValuePair<string, TableWriterStyle>(cell, cellStyle));
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
							string cell = childTableCellValueGetter.ValueFunc(child);
							cellsWithStyle.Add(new KeyValuePair<string, TableWriterStyle>(cell, cellStyle));
						}
					}

					int difference = (maximums[counter3] - children.Count()) * childAggregatedContainers.Count();

					for (int i = 0; i < difference; i++) {
						cellsWithStyle.Add(new KeyValuePair<string, TableWriterStyle>(null, null));
					}
					counter3++;
				}

				writer.WriteRow(cellsWithStyle);
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
		public void WriteRow(List<KeyValuePair<string, TableWriterStyle>> cells) {
			int cellsCount = cells.Count() - 1;
			int i = 0;
			foreach (KeyValuePair<string, TableWriterStyle> cell in cells) {
				if (cell.Key != null)
					_stringBuilder.Append(cell.Key);

				if (i < cellsCount)
					_stringBuilder.Append(_delimeter);
				i++;
			}
			_stringBuilder.AppendLine();
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
		private readonly HSSFWorkbook _workbook = new HSSFWorkbook();
		private readonly ISheet _workSheet;
		private readonly TableWriterStyle _style;

		public ExcelTableWriterSimple(TableWriterStyle style) {
			_style = style;
			_workSheet = _workbook.CreateSheet();
		}

		public void WriteHeader(params string[] cells) {
			IRow row = _workSheet.CreateRow(_rowIndex);
			row.Height = _style.HeaderHeight;

			ICellStyle cellStyle = ConvertToNpoiStyle(_style.HeaderCell);
			cellStyle.VerticalAlignment = VerticalAlignment.CENTER;

			int columnIndex = 0;
			foreach (string cell in cells) {
				ICell newCell = row.CreateCell(columnIndex);
				newCell.SetCellValue(cell);
				newCell.CellStyle = cellStyle;
				columnIndex++;
			}
			_rowIndex++;
		}

		public void WriteRow(List<KeyValuePair<string, TableWriterStyle>> cells) {
			IRow row = _workSheet.CreateRow(_rowIndex);
			ICellStyle cellStyle = ConvertToNpoiStyle(_style.RegularCell);

			int columnIndex = 0;
			foreach (KeyValuePair<string, TableWriterStyle> cell in cells) {
				ICell newCell = row.CreateCell(columnIndex);

				if (cell.Key != null)
					newCell.SetCellValue(cell.Key);

				if (cell.Value != null) {
					ICellStyle customCellStyle = ConvertToNpoiStyle(cell.Value.RegularCell);
					newCell.CellStyle = customCellStyle;
				}
				else {
					newCell.CellStyle = cellStyle;
				}
				columnIndex++;
			}
			_rowIndex++;
		}

		private ICellStyle ConvertToNpoiStyle(StyleSettings styleSettings) {
			IFont cellFont = _workbook.CreateFont();

			cellFont.FontName = styleSettings.FontName;
			cellFont.FontHeightInPoints = styleSettings.FontHeightInPoints;
			cellFont.IsItalic = styleSettings.Italic;
			cellFont.Underline = styleSettings.Underline ? FontUnderline.SINGLE.ByteValue : FontUnderline.NONE.ByteValue;
			cellFont.Boldweight = (short)styleSettings.BoldWeight;

			HSSFPalette palette = _workbook.GetCustomPalette();
			HSSFColor similarColor = palette.FindSimilarColor(styleSettings.FontColor.Red, styleSettings.FontColor.Green, styleSettings.FontColor.Blue);
			cellFont.Color = similarColor.GetIndex();

			ICellStyle cellStyle = _workbook.CreateCellStyle();
			cellStyle.SetFont(cellFont);

			if (styleSettings.BackgroundColor != null) {
				similarColor = palette.FindSimilarColor(styleSettings.BackgroundColor.Red, styleSettings.BackgroundColor.Green, styleSettings.BackgroundColor.Blue);
				cellStyle.FillForegroundColor = similarColor.GetIndex();
				cellStyle.FillPattern = FillPatternType.SOLID_FOREGROUND;
			}
			return cellStyle;
		}

		public void AutosizeColumns() {
			var columnLengths = new List<int>();

			for (int columnNum = 0; columnNum < _workSheet.GetRow(0).LastCellNum; columnNum++) {
				int columnMaximumLength = 0;
				for (int rowNum = 0; rowNum <= _workSheet.LastRowNum; rowNum++) {
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
