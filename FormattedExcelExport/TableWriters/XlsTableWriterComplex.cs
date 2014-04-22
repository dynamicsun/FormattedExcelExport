using System.Collections.Generic;
using System.Linq;
using FormattedExcelExport.Style;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;


namespace FormattedExcelExport.TableWriters {
	public sealed class XlsTableWriterComplex : XlsTableWriterBase, ITableWriterComplex {
		private byte _colorIndex;

		public XlsTableWriterComplex(TableWriterStyle style) : base(style) { }
		public void WriteHeader(params string[] cells) {
			IRow row = WorkSheet.CreateRow(RowIndex);
			row.Height = Style.HeaderHeight;

			ICellStyle cellStyle = ConvertToNpoiStyle(Style.HeaderCell);
			cellStyle.VerticalAlignment = VerticalAlignment.Center;

			int columnIndex = 0;
			foreach (string cell in cells) {
				ICell newCell = row.CreateCell(columnIndex);
				newCell.SetCellValue(cell);
				newCell.CellStyle = cellStyle;
				columnIndex++;
			}
			RowIndex++;
			_colorIndex = 0;
		}
		public void WriteRow(IEnumerable<KeyValuePair<string, TableWriterStyle>> cells, bool prependDelimeter = false) {
			IRow row = WorkSheet.CreateRow(RowIndex);
			ICellStyle cellStyle = ConvertToNpoiStyle(Style.RegularCell);

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
			RowIndex++;
		}
		public void WriteChildHeader(params string[] cells) {
			IRow row = WorkSheet.CreateRow(RowIndex);
			int columnIndex = 0;
			List<string> cellsList = cells.ToList();
			ICellStyle cellStyle = ConvertToNpoiStyle(Style.HeaderChildCell);

			if (_colorIndex >= Style.ColorsCollection.Count)
				_colorIndex = 0;

			AdHocCellStyle.Color color = Style.ColorsCollection.ElementAt(_colorIndex);
			if (color != null) {
				HSSFPalette palette = Workbook.GetCustomPalette();
				HSSFColor similarColor = palette.FindSimilarColor(color.Red, color.Green, color.Blue);
				cellStyle.FillForegroundColor = similarColor.GetIndex();
				cellStyle.FillPattern = FillPattern.SolidForeground;
				_colorIndex++;
			}

			foreach (string cell in cellsList) {
				ICell newCell = row.CreateCell(columnIndex);
				newCell.SetCellValue(cell);
				newCell.CellStyle = cellStyle;
				columnIndex++;
			}
			RowIndex++;
		}
		public void WriteChildRow(IEnumerable<KeyValuePair<string, TableWriterStyle>> cells, bool prependDelimeter = false) {
			IRow row = WorkSheet.CreateRow(RowIndex);

			ICellStyle cellStyle = ConvertToNpoiStyle(Style.RegularChildCell);

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
			RowIndex++;
		}
	}
}