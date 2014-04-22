using System.Collections.Generic;
using FormattedExcelExport.Style;
using NPOI.SS.UserModel;


namespace FormattedExcelExport.TableWriters {
	public sealed class XlsTableWriterSimple : XlsTableWriterBase, ITableWriterSimple {
		public XlsTableWriterSimple(TableWriterStyle style) : base(style) { }

		public void WriteHeader(List<string> cells) {
			IRow row = WorkSheet.CreateRow(RowIndex);
			row.Height = Style.HeaderHeight;

			ICellStyle cellStyle = ConvertToNpoiStyle(Style.HeaderCell);
			cellStyle.VerticalAlignment = VerticalAlignment.Top;

			int columnIndex = 0;
			foreach (string cell in cells) {
				ICell newCell = row.CreateCell(columnIndex);
				newCell.SetCellValue(cell);
				newCell.CellStyle = cellStyle;
				columnIndex++;
			}
			RowIndex++;
		}
		public void WriteRow(List<KeyValuePair<string, TableWriterStyle>> cells) {
			IRow row = WorkSheet.CreateRow(RowIndex);
			ICellStyle cellStyle = ConvertToNpoiStyle(Style.RegularCell);

			int columnIndex = 0;
			foreach (KeyValuePair<string, TableWriterStyle> cell in cells) {
				ICell newCell = row.CreateCell(columnIndex);

				if (cell.Key != null)
					newCell.SetCellValue(cell.Key);

				if (cell.Value != null) {
					if (cell.Value.RegularCell.BackgroundColor != null) {
						ICellStyle customCellStyle = ConvertToNpoiStyle(cell.Value.RegularCell);
						newCell.CellStyle = customCellStyle;
					}
					if (cell.Value.RegularChildCell.BackgroundColor != null) {
						ICellStyle customCellStyle = ConvertToNpoiStyle(cell.Value.RegularChildCell);
						newCell.CellStyle = customCellStyle;
					}
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