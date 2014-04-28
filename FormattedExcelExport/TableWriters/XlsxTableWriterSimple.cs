using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading;
using FormattedExcelExport.Style;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;


namespace FormattedExcelExport.TableWriters {
	public sealed class XlsxTableWriterSimple : XlsxTableWriterBase, ITableWriterSimple {
		public XlsxTableWriterSimple(TableWriterStyle style) : base(style) {
		}
		public void WriteHeader(List<string> cells) {
			IRow row = WorkSheet.CreateRow(RowIndex);
			row.Height = Style.HeaderHeight;

			ICellStyle cellStyle = ConvertToNpoiStyle(Style.HeaderCell);
			cellStyle.FillPattern = FillPattern.SolidForeground;
			cellStyle.VerticalAlignment = VerticalAlignment.Center;
			
			int columnIndex = 0;
			foreach (string cell in cells) {
				ICell newCell = row.CreateCell(columnIndex);
				newCell.SetCellValue(cell);
				newCell.CellStyle = cellStyle;

				XSSFCellStyle cs = (XSSFCellStyle)newCell.CellStyle;
				cs.SetFillForegroundColor(new XSSFColor(new[] { Style.HeaderCell.BackgroundColor.Red, Style.HeaderCell.BackgroundColor.Green, Style.HeaderCell.BackgroundColor.Blue }));
				
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
					ICellStyle customCellStyle = ConvertToNpoiStyle(cell.Value.RegularCell);
					newCell.CellStyle = customCellStyle;
					if (cell.Value.RegularCell.BackgroundColor != null) {
						XSSFCellStyle cs = (XSSFCellStyle) newCell.CellStyle;
						cs.SetFillForegroundColor(new XSSFColor(new[] { cell.Value.RegularCell.BackgroundColor.Red, cell.Value.RegularCell.BackgroundColor.Green, cell.Value.RegularCell.BackgroundColor.Blue }));
						cs.FillPattern = FillPattern.SolidForeground;
					}
					if (cell.Value.RegularChildCell.BackgroundColor != null) {
						XSSFCellStyle cs = (XSSFCellStyle)newCell.CellStyle;
						cs.SetFillForegroundColor(new XSSFColor(new[] { cell.Value.RegularChildCell.BackgroundColor.Red, cell.Value.RegularChildCell.BackgroundColor.Green, cell.Value.RegularChildCell.BackgroundColor.Blue }));
						cs.FillPattern = FillPattern.SolidForeground;
					}
				}
				else {
					newCell.CellStyle = cellStyle;
				}

				columnIndex++;
			}
			RowIndex++;
		}

		public MemoryStream GetStream() {
			MemoryStream memoryStream = new MemoryStream();
//			Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

			FileStream sw = File.Create("TestSimple.xlsx");
			Workbook.Write(sw);


			//			Workbook.Write(memoryStream);
			//			memoryStream.Position = 0;
			return memoryStream;
		}
	}
}
