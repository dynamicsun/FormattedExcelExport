using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;


namespace FormattedExcelExport {
	public interface ITableWriterComplex {
		void WriteHeader(params string[] cells);
		void WriteRow(IEnumerable<KeyValuePair<string, TableWriterStyle>> cells, bool prependDelimeter = false);
		void WriteChildHeader(params string[] cells);
		void WriteChildRow(IEnumerable<KeyValuePair<string, TableWriterStyle>> cells, bool prependDelimeter = false);
		void AutosizeColumns();
		MemoryStream GetStream();
	}

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

	public sealed class CsvTableWriterComplex : ITableWriterComplex {
		private readonly StringBuilder _stringBuilder = new StringBuilder();
		private readonly string _delimeter;
		public CsvTableWriterComplex(string delimeter = "\t") {
			_delimeter = delimeter;
		}

		public void WriteHeader(params string[] cells) {
			WriteRow(false, cells);
		}
		public void WriteRow(IEnumerable<KeyValuePair<string, TableWriterStyle>> cells, bool prependDelimeter = false) {
			int cellsCount = cells.Count() - 1;
			int i = 0;
			if (prependDelimeter) _stringBuilder.Append(_delimeter);
			foreach (KeyValuePair<string, TableWriterStyle> cell in cells) {
				_stringBuilder.Append(cell.Key);

				if (i < cellsCount)
					_stringBuilder.Append(_delimeter);
				i++;
			}
			_stringBuilder.AppendLine();
		}
		public void WriteRow(bool prependDelimeter, params string[] cells) {
			int cellsCount = cells.Length - 1;
			int i = 0;
			if (prependDelimeter) _stringBuilder.Append(_delimeter);
			foreach (string cell in cells) {
				_stringBuilder.Append(cell);

				if (i < cellsCount)
					_stringBuilder.Append(_delimeter);
				i++;
			}
			_stringBuilder.AppendLine();
		}
		public void WriteChildHeader(params string[] cells) {
			WriteHeader(cells);
		}
		public void WriteChildRow(IEnumerable<KeyValuePair<string, TableWriterStyle>> cells, bool prependDelimeter = false) {
			WriteRow(cells, prependDelimeter);
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

	public sealed class ExcelTableWriterComplex : ITableWriterComplex {
		private int _rowIndex;
		private readonly HSSFWorkbook _workbook = new HSSFWorkbook();
		private readonly ISheet _workSheet;
		private readonly TableWriterStyle _style;
		private byte _colorIndex;

		public ExcelTableWriterComplex(TableWriterStyle style) {
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
			_colorIndex = 0;
		}
		public void WriteRow(IEnumerable<KeyValuePair<string, TableWriterStyle>> cells, bool prependDelimeter = false) {
			IRow row = _workSheet.CreateRow(_rowIndex);

			ICellStyle cellStyle = ConvertToNpoiStyle(_style.RegularCell);

			int columnIndex = 0;
			if (prependDelimeter) {
				ICell newCell = row.CreateCell(columnIndex);
				newCell.SetCellValue("");
				newCell.CellStyle = cellStyle;
				
				columnIndex++;
			}
			foreach (KeyValuePair<string, TableWriterStyle> cell in cells) {
				ICell newCell = row.CreateCell(columnIndex);
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
		public void WriteChildHeader(params string[] cells) {
			IRow row = _workSheet.CreateRow(_rowIndex);
			int columnIndex = 0;
			List<string> cellsList = cells.ToList();

			ICellStyle cellStyle = ConvertToNpoiStyle(_style.HeaderChildCell);

			if (_colorIndex >= _style.ColorsCollection.Count)
				_colorIndex = 0;

			StyleSettings.Color color = _style.ColorsCollection.ElementAt(_colorIndex);
			if (color != null) {
				HSSFPalette palette = _workbook.GetCustomPalette();
				HSSFColor similarColor = palette.FindSimilarColor(color.Red, color.Green, color.Blue);
				cellStyle.FillForegroundColor = similarColor.GetIndex();
				cellStyle.FillPattern = FillPatternType.SOLID_FOREGROUND;
				_colorIndex++;
			}
			
			foreach (string cell in cellsList) {
				ICell newCell = row.CreateCell(columnIndex);
				newCell.SetCellValue(cell);
				newCell.CellStyle = cellStyle;
				columnIndex++;
			}
			_rowIndex++;
		}
		public void WriteChildRow(IEnumerable<KeyValuePair<string, TableWriterStyle>> cells, bool prependDelimeter = false) {
			IRow row = _workSheet.CreateRow(_rowIndex);

			ICellStyle cellStyle = ConvertToNpoiStyle(_style.RegularChildCell);

			int columnIndex = 0;
			if (prependDelimeter) {
				ICell newCell = row.CreateCell(columnIndex);
				newCell.SetCellValue("");
				newCell.CellStyle = cellStyle;

				columnIndex++;
			}
			foreach (KeyValuePair<string, TableWriterStyle> cell in cells) {
				ICell newCell = row.CreateCell(columnIndex);
				newCell.SetCellValue(cell.Key);

				if (cell.Value != null) {
					ICellStyle customCellStyle = ConvertToNpoiStyle(cell.Value.RegularChildCell);
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
		public void AutosizeColumns() { // Set columns width based on the contents width of their corresponding header cells
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
