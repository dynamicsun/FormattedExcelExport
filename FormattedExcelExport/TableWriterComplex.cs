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
		void WriteRow(bool prependDelimeter = false, params string[] cells);
		void WriteChildHeader(params string[] cells);
		void AutosizeColumns();
		MemoryStream GetStream();
	}

	public static class TableWriterComplex {
		public static MemoryStream Write<TModel>(ITableWriterComplex writer, IEnumerable<TModel> models, TableConfiguration parentTableConfiguration) {
			List<string> headerNamesList = parentTableConfiguration.ColumnsMap.Keys.ToList();
			headerNamesList.Insert(0, parentTableConfiguration.Title);
			Func<object, string>[] parentTableCellValueGetters = parentTableConfiguration.ColumnsMap.Values.ToArray();
			List<ChildTableConfiguration> childTableConfigurations = parentTableConfiguration.ChildrenMap;

			foreach (TModel model in models) {
				writer.WriteHeader(headerNamesList.ToArray());
				var cells = new List<string>();
				foreach (Func<object, string> parentTableCellValueGetter in parentTableCellValueGetters) {
					string cell = parentTableCellValueGetter(model);
					cells.Add(cell);
				}
				writer.WriteRow(true, cells.ToArray());

				foreach (ChildTableConfiguration childTableConfiguration in childTableConfigurations) {
					IEnumerable<object> children = childTableConfiguration.Getter(model);
					Func<object, string>[] childTableCellValueGetters = childTableConfiguration.ColumnsMap.Values.ToArray();
					List<string> childHeaderNamesList = childTableConfiguration.ColumnsMap.Keys.ToList();
					childHeaderNamesList.Insert(0, childTableConfiguration.Title);
					writer.WriteChildHeader(childHeaderNamesList.ToArray());

					foreach (object child in children) {
						var childCells = new List<string>();
						foreach (Func<object, string> childTableCellValueGetter in childTableCellValueGetters) {
							string cell = childTableCellValueGetter(child);
							childCells.Add(cell);
						}
						writer.WriteRow(true, childCells.ToArray());
					}
				}
				writer.WriteRow();
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
		private bool _isFirstChildHeader = true;

		public ExcelTableWriterComplex(TableWriterStyle style) {
			_style = style;
			_workSheet = _workbook.CreateSheet();
		}
		public void WriteHeader(params string[] cells) {
			IRow row = _workSheet.CreateRow(_rowIndex);
			row.Height = _style.HeaderHeight;


			IFont cellFont = _workbook.CreateFont();
			cellFont.FontName = _style.HeaderCell.FontName;
			cellFont.FontHeightInPoints = _style.HeaderCell.FontHeightInPoints;
			cellFont.IsItalic = _style.HeaderCell.Italic;
			cellFont.Underline = _style.HeaderCell.Underline ? FontUnderline.SINGLE.ByteValue : FontUnderline.NONE.ByteValue;
			cellFont.Boldweight = (short)_style.HeaderCell.BoldWeight;

			HSSFPalette palette = _workbook.GetCustomPalette();
			palette.SetColorAtIndex(HSSFColor.LIGHT_ORANGE.index, _style.HeaderCell.FontColor.Red, _style.HeaderCell.FontColor.Green, _style.HeaderCell.FontColor.Blue); // Redefine some arbitrary color so we could create our custom color
			cellFont.Color = HSSFColor.LIGHT_ORANGE.index;

			ICellStyle cellStyle = _workbook.CreateCellStyle();
			cellStyle.SetFont(cellFont);

			palette.SetColorAtIndex(HSSFColor.AQUA.index, _style.HeaderCell.BackgroundColor.Red, _style.HeaderCell.BackgroundColor.Green, _style.HeaderCell.BackgroundColor.Blue); // Redefine some arbitrary color so we could create our custom color
			cellStyle.FillForegroundColor = HSSFColor.AQUA.index;
			cellStyle.FillPattern = FillPatternType.SOLID_FOREGROUND;

			cellStyle.VerticalAlignment = VerticalAlignment.CENTER;

			int columnIndex = 0;
			foreach (string cell in cells) {
				ICell newCell = row.CreateCell(columnIndex);
				newCell.SetCellValue(cell);
				newCell.CellStyle = cellStyle;
				columnIndex++;
			}
			_rowIndex++;
			_isFirstChildHeader = true;
		}
		public void WriteRow(bool prependDelimeter = false, params string[] cells) {
			IRow newRow = _workSheet.CreateRow(_rowIndex);
			int columnIndex = 0;
			if (prependDelimeter) {
				newRow.CreateCell(columnIndex).SetCellValue("");
				columnIndex++;
			}
			foreach (var cell in cells) {
				newRow.CreateCell(columnIndex).SetCellValue(cell);
				columnIndex++;
			}
			_rowIndex++;
		}
		public void WriteChildHeader(params string[] cells) {
			IRow newRow = _workSheet.CreateRow(_rowIndex);
			int columnIndex = 0;
			List<string> cellsList = cells.ToList();

			if (cellsList.Any()) {
				ICell cell = newRow.CreateCell(columnIndex);
				cell.SetCellValue(cellsList.ElementAt(0));

				if (_isFirstChildHeader) {
//					cell.CellStyle = _style.ChildHeaderCellStylePremier;
					_isFirstChildHeader = false;
				}
				else {
//					cell.CellStyle = _style.ChildHeaderCellStyleNext;
				}

				cellsList.RemoveAt(0);
				columnIndex++;
			}

			foreach (string cell in cellsList) {
				newRow.CreateCell(columnIndex).SetCellValue(cell);
				columnIndex++;
			}
			_rowIndex++;
		}
		public void AutosizeColumns() { // Set columns width based on the contents width of their corresponding header cells
			var columnLengths = new List<int>();

			for (int columnNum = 0; columnNum < _workSheet.GetRow(0).LastCellNum; columnNum++) { // Looping through the each cell of the first row
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
