using System.Collections.Generic;
using System.IO;
using System.Linq;
using FormattedExcelExport.Style;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;


namespace FormattedExcelExport.TableWriters {
	public abstract class XlsTableWriterBase {
        protected const int MaxRowIndex = 65535;
	    protected const int MaxWidth = 5000;
		protected readonly HSSFWorkbook Workbook;
		protected ISheet WorkSheet;
		protected readonly TableWriterStyle Style;
        protected ICellStyle CellStyle;
        protected ICellStyle HeaderCellStyle;
	    protected ICellStyle DateCellStyle;
		protected XlsTableWriterBase(TableWriterStyle style) {
			Workbook = new HSSFWorkbook();
			Style = style ?? new TableWriterStyle();
			WorkSheet = Workbook.CreateSheet();
            CellStyle = ConvertToNpoiStyle(Style.RegularCell);
            HeaderCellStyle = ConvertToNpoiStyle(Style.HeaderCell);
            var dataFormat = "m/d/yy";
		    short dataFormatValue;
            var builtinFormatId = HSSFDataFormat.GetBuiltinFormat(dataFormat);
            if (builtinFormatId != -1)
                dataFormatValue = builtinFormatId;
            else {
                var dataFormatCustom = Workbook.CreateDataFormat();
                dataFormatValue = dataFormatCustom.GetFormat(dataFormat);
            }
		    DateCellStyle = ConvertToNpoiStyle(Style.RegularCell, dataFormatValue);
		}

		public void AutosizeColumns() {
		    for (int sheetNumber = 0; sheetNumber < Workbook.NumberOfSheets; sheetNumber++) {
		        var columnLengths = new List<int>();
		        WorkSheet = Workbook.GetSheetAt(sheetNumber);
		        for (int columnNum = 0; columnNum < WorkSheet.GetRow(0).LastCellNum; columnNum++) {
		            int columnMaximumLength = 0;
		            for (int rowNum = 0; rowNum <= WorkSheet.LastRowNum; rowNum++) {
		                IRow currentRow = WorkSheet.GetRow(rowNum);

		                if (!currentRow.Cells.Any()) continue;
		                ICell cell = currentRow.GetCell(columnNum);
		                if (cell == null) continue;
                        if (cell.CellType == CellType.String) 
                            if (cell.StringCellValue.Length > columnMaximumLength) 
                                columnMaximumLength = cell.StringCellValue.Length;
                        if (cell.CellType == CellType.Numeric)
                            if (cell.NumericCellValue.ToString().Length > columnMaximumLength)
                                columnMaximumLength = cell.NumericCellValue.ToString().Length;
		            }
		            columnLengths.Add(columnMaximumLength);
		        }

		        for (int i = 0; i < WorkSheet.GetRow(0).LastCellNum; i++) {
		            int width = columnLengths.ElementAt(i)*Style.FontFactor + Style.FontAbsoluteTerm;
                    WorkSheet.SetColumnWidth(i, width < MaxWidth ? width : MaxWidth);
		            for (int j = 0; j < WorkSheet.LastRowNum; j++) {
		                if (WorkSheet.GetRow(j).GetCell(i) != null) {
		                    WorkSheet.GetRow(j).GetCell(i).CellStyle.WrapText = true;
		                }
		            }
		        }
		    }
		}

		protected ICellStyle ConvertToNpoiStyle(AdHocCellStyle adHocCellStyle, short dateFormat = 0) {
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
		    cellStyle.DataFormat = dateFormat;
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