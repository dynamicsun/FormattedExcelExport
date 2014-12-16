using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using FormattedExcelExport.Style;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace FormattedExcelExport.TableWriters {
    public abstract class XlsxTableWriterBase {
        protected const int MaxRowIndex = 1048575;
        protected const int MaxWidth = 5000;
        protected const int MinWidth = 2000;
        protected readonly ExcelPackage Package;
        protected ExcelWorksheet WorkSheet;
        protected readonly TableWriterStyle Style;
        protected int SheetNumber;
        protected List<string> LastHeader;

        protected XlsxTableWriterBase(TableWriterStyle style) {
            Package = new ExcelPackage();
            Style = style;
            SheetNumber = 0;
            WorkSheet = Package.Workbook.Worksheets.Add("Sheet " + SheetNumber);
        }

        public void AutosizeColumns() {
            const int conversionFactorWidth = 256;
            foreach (var sheet in Package.Workbook.Worksheets) {
                var columnLengths = new List<int>();
                for (var columnNum = 1; columnNum <= sheet.Dimension.End.Column; columnNum++) {
                    var columnMaxixumLength = 0;
                    for (var rowNum = 1; rowNum <= sheet.Dimension.End.Row; rowNum++) {
                        if (WorkSheet.Cells[rowNum, columnNum] == null) continue;

                        if (WorkSheet.Cells[rowNum, columnNum].Value != null && WorkSheet.Cells[rowNum, columnNum].Value.ToString().Length > columnMaxixumLength) {
                            columnMaxixumLength = WorkSheet.Cells[rowNum, columnNum].Value.ToString().Length;
                        }
                    }
                    columnLengths.Add(columnMaxixumLength);
                }
                for (var i = 1; i <= sheet.Dimension.End.Column; i++) {
                    var a = sheet.Cells[1, i].Value;
                    var width = columnLengths.ElementAt(i - 1)*Style.FontFactor + Style.FontAbsoluteTerm;
                    sheet.Column(i).Width = (width < MaxWidth ? width : MaxWidth)/conversionFactorWidth;
                    sheet.Column(i).Style.WrapText = true;
                }
            }
        }

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
        public MemoryStream GetStream() {
            var memoryStream = new MemoryStream();
            Package.SaveAs(memoryStream);
            memoryStream.Position = 0;
            return memoryStream;
        }

        protected void WriteHeaderBase(IEnumerable<string> cells, int rowIndex) {
            var font = ConvertCellStyle(Style.HeaderCell);
            var columnIndex = 1;
            foreach (var cell in cells) {
                var newCell = WorkSheet.Cells[rowIndex, columnIndex];
                newCell.Value = cell;
                newCell.Style.Font.SetFromFont(font);
                newCell.Style.Font.Color.SetColor(Color.FromArgb(Style.HeaderCell.FontColor.Red, Style.HeaderCell.FontColor.Green, Style.HeaderCell.FontColor.Blue));
                newCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                newCell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(Style.HeaderCell.BackgroundColor.Red, Style.HeaderCell.BackgroundColor.Green, Style.HeaderCell.BackgroundColor.Blue));
                newCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                columnIndex++;
            }
            LastHeader = cells.ToList();
        }
        protected void WriteRowBase(IEnumerable<KeyValuePair<dynamic, TableWriterStyle>> cells, int rowIndex, bool prependDelimeter = false, bool isChildRow = false) {
            var columnIndex = 1;
            var defaultStyle = isChildRow ? Style.RegularChildCell : Style.RegularCell;
            var font = ConvertCellStyle(defaultStyle);
            if (prependDelimeter) {
                var newCell = WorkSheet.Cells[rowIndex, columnIndex];
                newCell.Value = "";
                if (defaultStyle.BackgroundColor != null) {
                    newCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    newCell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(defaultStyle.BackgroundColor.Red, defaultStyle.BackgroundColor.Green, defaultStyle.BackgroundColor.Blue));
                }
                columnIndex++;
            }
            foreach (var cell in cells) {
                var newCell = WorkSheet.Cells[rowIndex, columnIndex];
                if (cell.Key != null) if (cell.Key is DateTime?) {
                        var date = (DateTime)cell.Key;
                        newCell.Formula = "=Date(" + date.Year + "," + date.Month + "," + date.Day + ")";
                        newCell.Style.Numberformat.Format = "dd.mm.yyyy";
                    } else newCell.Value = cell.Key;
                if (cell.Value != null) {
                    var customStyle = isChildRow ? cell.Value.RegularChildCell : cell.Value.RegularCell;
                    font = ConvertCellStyle(customStyle);
                    newCell.Style.Font.SetFromFont(font);
                    newCell.Style.Font.Color.SetColor(Color.FromArgb(Style.HeaderCell.FontColor.Red, Style.HeaderCell.FontColor.Green, Style.HeaderCell.FontColor.Blue));
                    if (customStyle.BackgroundColor != null) {
                        newCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        newCell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(customStyle.BackgroundColor.Red, customStyle.BackgroundColor.Green, customStyle.BackgroundColor.Blue));
                    }
                    newCell.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                } else {
                    newCell.Style.Font.SetFromFont(font);
                    newCell.Style.Font.Color.SetColor(Color.FromArgb(defaultStyle.FontColor.Red, defaultStyle.FontColor.Green, defaultStyle.FontColor.Blue));
                    newCell.Style.Fill.PatternType = ExcelFillStyle.None;
                    newCell.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                }
                columnIndex++;
            }
        }
    }
}
