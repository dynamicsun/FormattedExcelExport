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
            
        }

        //protected ExcelStyle ConvertToEPPlusStyle(AdHocCellStyle adHocCellStyle) {
        //    ExcelFont cellFont = WorkSheet.Cells[1, 1].Style.Font;

        //    cellFont.Name = adHocCellStyle.FontName;
        //    cellFont.Italic = adHocCellStyle.Italic;
        //    if (adHocCellStyle.Underline) {
        //        cellFont.UnderLineType = ExcelUnderLineType.Single;
        //    }
        //    if ((short)adHocCellStyle.BoldWeight == 700) {
        //        cellFont.Bold = true;
        //    }
        //    cellFont.Color.SetColor(Color.FromArgb(adHocCellStyle.FontColor.Red, adHocCellStyle.FontColor.Green, adHocCellStyle.FontColor.Blue));

        //    ExcelStyle cellStyle = WorkSheet.Cells[1, 1].Style;
        //    cellStyle.Font = cellFont;
        //    return cellStyle;
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
