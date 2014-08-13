using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FormattedExcelExport.Style;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace FormattedExcelExport.TableWriters {
    public abstract class XlsxTableWriterBase {
        protected const int MaxRowIndex = 1048575;
        protected readonly ExcelPackage Package;
        protected ExcelWorksheet WorkSheet;
        protected readonly TableWriterStyle Style;
        protected int SheetNumber;

        protected XlsxTableWriterBase(TableWriterStyle style) {
            Package = new ExcelPackage();
            Style = style;
            SheetNumber = 0;
            WorkSheet = Package.Workbook.Worksheets.Add("Sheet " + SheetNumber);
        }

        public void AutosizeColumns() {
            foreach (var sheet in Package.Workbook.Worksheets) {
                List<int> columnLengths = new List<int>();
                for (int columnNum = 1; columnNum < sheet.Dimension.End.Column; columnNum++) {
                    int columnMaxixumLength = 0;
                    for (int rowNum = 1; rowNum < sheet.Dimension.End.Row; rowNum++) {
                        if (WorkSheet.Cells[rowNum, columnNum] == null) continue;

                        if (WorkSheet.Cells[rowNum, columnNum].Value != null && WorkSheet.Cells[rowNum, columnNum].Value.ToString().Length > columnMaxixumLength) {
                            columnMaxixumLength = WorkSheet.Cells[rowNum, columnNum].Value.ToString().Length;
                        }
                    }
                    columnLengths.Add(columnMaxixumLength);
                }

                for (int i = 1; i < sheet.Dimension.End.Column - 1; i++) {
                    int width = (columnLengths.ElementAt(i - 1) * Style.FontFactor + Style.FontAbsoluteTerm) / 256;
                    sheet.Column(i).Width = width < Style.MaxColumnWidth ? width : Style.MaxColumnWidth;
                }
            }
        }

        //public void AutosizeColumns() {
        //    for (int sheetNumber = 0; sheetNumber < Workbook.NumberOfSheets; sheetNumber++) {
        //        var columnLengths = new List<int>();
        //        WorkSheet = Workbook.GetSheetAt(sheetNumber);
        //        for (int columnNum = 0; columnNum < WorkSheet.GetRow(0).LastCellNum; columnNum++) {
        //            int columnMaximumLength = 0;
        //            for (int rowNum = 0; rowNum <= WorkSheet.LastRowNum; rowNum++) {
        //                IRow currentRow = WorkSheet.GetRow(rowNum);

        //                if (!currentRow.Cells.Any()) continue;
        //                ICell cell = currentRow.GetCell(columnNum);
        //                if (cell == null) continue;

        //                if (cell.StringCellValue.Length > columnMaximumLength)
        //                    columnMaximumLength = cell.StringCellValue.Length;
        //            }
        //            columnLengths.Add(columnMaximumLength);
        //        }

        //        for (int i = 0; i < WorkSheet.GetRow(0).LastCellNum; i++) {
        //            int width = columnLengths.ElementAt(i) * Style.FontFactor + Style.FontAbsoluteTerm;
        //            WorkSheet.SetColumnWidth(i, width < Style.MaxColumnWidth ? width : Style.MaxColumnWidth);
        //        }
        //    }
        //}

        protected Font ConvertCellStyle(AdHocCellStyle adHocCellStyle) {
            Font font = null;
            if (adHocCellStyle.Italic && adHocCellStyle.BoldWeight == AdHocCellStyle.FontBoldWeight.Normal && !adHocCellStyle.Underline) {
                font = new Font(adHocCellStyle.FontName, adHocCellStyle.FontHeightInPoints, FontStyle.Italic);
            }
            if (!adHocCellStyle.Italic && adHocCellStyle.BoldWeight == AdHocCellStyle.FontBoldWeight.Bold && !adHocCellStyle.Underline) {
                font = new Font(adHocCellStyle.FontName, adHocCellStyle.FontHeightInPoints, FontStyle.Bold);
            }
            if (!adHocCellStyle.Italic && adHocCellStyle.BoldWeight == AdHocCellStyle.FontBoldWeight.Normal && adHocCellStyle.Underline) {
                font = new Font(adHocCellStyle.FontName, adHocCellStyle.FontHeightInPoints, FontStyle.Underline);
            }

            if (adHocCellStyle.Italic && adHocCellStyle.BoldWeight == AdHocCellStyle.FontBoldWeight.Normal && adHocCellStyle.Underline) {
                font = new Font(adHocCellStyle.FontName, adHocCellStyle.FontHeightInPoints, FontStyle.Italic & FontStyle.Underline);
            }
            if (adHocCellStyle.Italic && adHocCellStyle.BoldWeight == AdHocCellStyle.FontBoldWeight.Bold && !adHocCellStyle.Underline) {
                font = new Font(adHocCellStyle.FontName, adHocCellStyle.FontHeightInPoints, FontStyle.Italic & FontStyle.Bold);
            }
            if (!adHocCellStyle.Italic && adHocCellStyle.BoldWeight == AdHocCellStyle.FontBoldWeight.Bold && adHocCellStyle.Underline) {
                font = new Font(adHocCellStyle.FontName, adHocCellStyle.FontHeightInPoints, FontStyle.Bold & FontStyle.Underline);
            }

            if (adHocCellStyle.Italic && adHocCellStyle.BoldWeight == AdHocCellStyle.FontBoldWeight.Bold && adHocCellStyle.Underline) {
                font = new Font(adHocCellStyle.FontName, adHocCellStyle.FontHeightInPoints, FontStyle.Italic & FontStyle.Bold & FontStyle.Underline);
            }

            if (!adHocCellStyle.Italic && adHocCellStyle.BoldWeight == AdHocCellStyle.FontBoldWeight.Normal && !adHocCellStyle.Underline) {
                font = new Font(adHocCellStyle.FontName, adHocCellStyle.FontHeightInPoints);
            }
            return font;
        }
    }
}
