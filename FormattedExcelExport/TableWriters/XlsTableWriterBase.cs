using System.Collections.Generic;
using System.IO;
using System.Linq;
using FormattedExcelExport.Style;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;


namespace FormattedExcelExport.TableWriters {
	public abstract class XlsTableWriterBase {
		protected int RowIndex;
		protected readonly HSSFWorkbook Workbook;
		protected readonly ISheet WorkSheet;
		protected readonly TableWriterStyle Style;

		protected XlsTableWriterBase(TableWriterStyle style) {
			Workbook = new HSSFWorkbook();
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
			cellFont.Underline = adHocCellStyle.Underline ? FontUnderlineType.Single : FontUnderlineType.None;
			cellFont.Boldweight = (short)adHocCellStyle.BoldWeight;

			HSSFPalette palette = Workbook.GetCustomPalette();
			HSSFColor similarColor = palette.FindSimilarColor(adHocCellStyle.FontColor.Red, adHocCellStyle.FontColor.Green, adHocCellStyle.FontColor.Blue);
			cellFont.Color = similarColor.GetIndex();

			ICellStyle cellStyle = Workbook.CreateCellStyle();
			cellStyle.SetFont(cellFont);

			if (adHocCellStyle.BackgroundColor != null) {
				similarColor = palette.FindSimilarColor(adHocCellStyle.BackgroundColor.Red, adHocCellStyle.BackgroundColor.Green, adHocCellStyle.BackgroundColor.Blue);
				cellStyle.FillForegroundColor = similarColor.GetIndex();
				cellStyle.FillPattern = FillPattern.SolidForeground;
			}
			return cellStyle;
		}

		public MemoryStream GetStream() {
			MemoryStream memoryStream = new MemoryStream();

			Workbook.Write(memoryStream);
			memoryStream.Position = 0;
			return memoryStream;
		}
	}
}