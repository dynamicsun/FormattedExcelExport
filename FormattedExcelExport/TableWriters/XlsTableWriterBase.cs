using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FormattedExcelExport.Style;
using NPOI.HSSF.UserModel;
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
		    for (var sheetNumber = 0; sheetNumber < Workbook.NumberOfSheets; sheetNumber++) {
		        var columnLengths = new List<int>();
		        WorkSheet = Workbook.GetSheetAt(sheetNumber);
		        for (var columnNum = 0; columnNum < WorkSheet.GetRow(0).LastCellNum; columnNum++) {
		            var columnMaximumLength = 0;
		            for (var rowNum = 0; rowNum <= WorkSheet.LastRowNum; rowNum++) {
		                var currentRow = WorkSheet.GetRow(rowNum);

		                if (!currentRow.Cells.Any()) continue;
		                var cell = currentRow.GetCell(columnNum);
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

		        for (var i = 0; i < WorkSheet.GetRow(0).LastCellNum; i++) {
		            var width = columnLengths.ElementAt(i)*Style.FontFactor + Style.FontAbsoluteTerm;
                    WorkSheet.SetColumnWidth(i, width < MaxWidth ? width : MaxWidth);
		            for (var j = 0; j < WorkSheet.LastRowNum; j++) {
		                if (WorkSheet.GetRow(j).GetCell(i) != null) {
		                    WorkSheet.GetRow(j).GetCell(i).CellStyle.WrapText = true;
		                }
		            }
		        }
		    }
		}
        protected void WriteRowBase(IEnumerable<KeyValuePair<dynamic, TableWriterStyle>> cells, int rowIndex, bool prependDelimeter = false, bool isChildRow = false) {
            var row = WorkSheet.CreateRow(rowIndex);
            var columnIndex = 0;
            if (prependDelimeter) {
                var newCell = row.CreateCell(columnIndex);
                newCell.SetCellValue("");
                newCell.CellStyle = CellStyle;
                columnIndex++;
            }
            foreach (var cell in cells) {
                var newCell = row.CreateCell(columnIndex);
                newCell.SetCellValue(cell.Key ?? string.Empty);
                if (cell.Key != null && cell.Key is DateTime?) {
                    newCell.CellStyle = DateCellStyle;
                } else {
                    if (cell.Value != null) {
                        ICellStyle customCellStyle;
                        customCellStyle = ConvertToNpoiStyle(isChildRow ? cell.Value.RegularChildCell : cell.Value.RegularCell);
                        newCell.CellStyle = customCellStyle;
                    } else {
                        newCell.CellStyle = CellStyle;
                    }
                }
                columnIndex++;
            }
        }
		protected ICellStyle ConvertToNpoiStyle(AdHocCellStyle adHocCellStyle, short dateFormat = 0) {
			var cellFont = Workbook.CreateFont();

			cellFont.FontName = adHocCellStyle.FontName;
			cellFont.FontHeightInPoints = adHocCellStyle.FontHeightInPoints;
			cellFont.IsItalic = adHocCellStyle.Italic;
			cellFont.Underline = adHocCellStyle.Underline ? FontUnderlineType.Single : FontUnderlineType.None;
			cellFont.Boldweight = (short)adHocCellStyle.BoldWeight;

			var palette = Workbook.GetCustomPalette();
			var similarColor = palette.FindSimilarColor(adHocCellStyle.FontColor.Red, adHocCellStyle.FontColor.Green, adHocCellStyle.FontColor.Blue);
			cellFont.Color = similarColor.GetIndex();

			var cellStyle = Workbook.CreateCellStyle();
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
			var memoryStream = new MemoryStream();

			Workbook.Write(memoryStream);
			memoryStream.Position = 0;
			return memoryStream;
		}
	}
}