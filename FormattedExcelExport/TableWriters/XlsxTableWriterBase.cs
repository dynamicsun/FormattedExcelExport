using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using FormattedExcelExport.Style;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;


namespace FormattedExcelExport.TableWriters {
	public abstract class XlsxTableWriterBase {
		protected int RowIndex;
		protected readonly XSSFWorkbook Workbook;
		protected readonly ISheet WorkSheet;
		protected readonly TableWriterStyle Style;

		protected XlsxTableWriterBase(TableWriterStyle style) {
			Workbook = new XSSFWorkbook();
			Style = style;
			WorkSheet = Workbook.CreateSheet();
		}

		public void AutosizeColumns() {
			var columnLengths = new List<int>();

			for (int columnNum = 0; columnNum < WorkSheet.GetRow(0).LastCellNum; columnNum++) {
				int columnMaximumLength = 0;
				for (int rowNum = 0; rowNum <= WorkSheet.LastRowNum; rowNum++) {
					IRow currentRow = WorkSheet.GetRow(rowNum);

					if (!currentRow.Cells.Any()) continue;
					ICell cell = currentRow.GetCell(columnNum);
					if (cell == null) continue;

					if (cell.StringCellValue.Length > columnMaximumLength)
						columnMaximumLength = cell.StringCellValue.Length;
				}
				columnLengths.Add(columnMaximumLength);
			}

			for (int i = 0; i < WorkSheet.GetRow(0).LastCellNum; i++) {
				int width = columnLengths.ElementAt(i) * Style.FontFactor + Style.FontAbsoluteTerm;
				WorkSheet.SetColumnWidth(i, width < Style.MaxColumnWidth ? width : Style.MaxColumnWidth);
			}
		}

		protected ICellStyle ConvertToNpoiStyle(AdHocCellStyle adHocCellStyle) {
			IFont cellFont = Workbook.CreateFont();

			cellFont.FontName = adHocCellStyle.FontName;
			cellFont.FontHeightInPoints = adHocCellStyle.FontHeightInPoints;
			cellFont.IsItalic = adHocCellStyle.Italic;
			if (adHocCellStyle.Underline) {
				cellFont.Underline = FontUnderlineType.Single;
			}
			cellFont.Boldweight = (short)adHocCellStyle.BoldWeight;
			((XSSFFont)cellFont).SetColor(new XSSFColor(new[] { adHocCellStyle.FontColor.Red, adHocCellStyle.FontColor.Green, adHocCellStyle.FontColor.Blue }));

			ICellStyle cellStyle = Workbook.CreateCellStyle();
			cellStyle.SetFont(cellFont);

			return cellStyle;
		}
	}
}
