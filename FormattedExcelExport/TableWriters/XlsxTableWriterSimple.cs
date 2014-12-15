using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using FormattedExcelExport.Infrastructure;
using FormattedExcelExport.Style;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace FormattedExcelExport.TableWriters {
    public sealed class XlsxTableWriterSimple : XlsxTableWriterBase, ITableWriterSimple {
        public XlsxTableWriterSimple(TableWriterStyle style) : base(style) {}
        private int _rowIndex = 1;
        private List<string> _lastHeader;

        public int RowIndex {
            get { return _rowIndex; }
            set {
                if (_rowIndex < MaxRowIndex) {
                    _rowIndex = value;
                } else {
                    SheetNumber = SheetNumber + 1;
                    WorkSheet = Package.Workbook.Worksheets.Add("Sheet " + SheetNumber);
                    _rowIndex = 1;
                    WriteHeader(_lastHeader);
                }
            }
        }

        public void WriteHeader(List<string> cells) {
            ExcelRow row = WorkSheet.Row(RowIndex);
            //row.Height = Style.HeaderHeight / 20;
            Font font = ConvertCellStyle(Style.HeaderCell);
            int columnIndex = 1;
            foreach (string cell in cells) {
                ExcelRange newCell = WorkSheet.Cells[RowIndex, columnIndex];
                newCell.Value = cell;
                newCell.Style.Font.SetFromFont(font);
                newCell.Style.Font.Color.SetColor(Color.FromArgb(Style.HeaderCell.FontColor.Red, Style.HeaderCell.FontColor.Green, Style.HeaderCell.FontColor.Blue));
                newCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                newCell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(Style.HeaderCell.BackgroundColor.Red, Style.HeaderCell.BackgroundColor.Green, Style.HeaderCell.BackgroundColor.Blue));
                newCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                columnIndex++;
            }
            _lastHeader = cells;
            RowIndex++;
        }

        public void WriteRow(List<KeyValuePair<dynamic, TableWriterStyle>> cells) {
            int columnIndex = 1;
            Font font = ConvertCellStyle(Style.RegularCell);
            foreach (KeyValuePair<dynamic, TableWriterStyle> cell in cells) {
                ExcelRange newCell = WorkSheet.Cells[RowIndex, columnIndex];
                if (cell.Key != null)
                    if (cell.Key is DateTime?) {
                        var date = (DateTime) cell.Key;
                        newCell.Formula = "=Date(" + date.Year + "," + date.Month + "," + date.Day + ")";
                        newCell.Style.Numberformat.Format = "dd.mm.yyyy";
                    }
                    else newCell.Value = cell.Key;
                if (cell.Value != null) {
                    font = ConvertCellStyle(cell.Value.RegularCell);
                    newCell.Style.Font.SetFromFont(font);
                    newCell.Style.Font.Color.SetColor(Color.FromArgb(Style.HeaderCell.FontColor.Red, Style.HeaderCell.FontColor.Green, Style.HeaderCell.FontColor.Blue));
                    if (cell.Value.RegularCell.BackgroundColor != null) {
                        newCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        newCell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(cell.Value.RegularCell.BackgroundColor.Red, cell.Value.RegularCell.BackgroundColor.Green, cell.Value.RegularCell.BackgroundColor.Blue));
                    }
                    if (cell.Value.RegularChildCell.BackgroundColor != null) {
                        newCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        newCell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(cell.Value.RegularChildCell.BackgroundColor.Red, cell.Value.RegularChildCell.BackgroundColor.Green, cell.Value.RegularChildCell.BackgroundColor.Blue));
                    }
                    newCell.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                }
                else {
                    newCell.Style.Font.SetFromFont(font);
                    newCell.Style.Font.Color.SetColor(Color.FromArgb(Style.RegularCell.FontColor.Red, Style.RegularCell.FontColor.Green, Style.RegularCell.FontColor.Blue));
                    newCell.Style.Fill.PatternType = ExcelFillStyle.None;
                    newCell.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                }
                columnIndex++;
            }
            RowIndex++;
        }

        public MemoryStream GetStream() {
            MemoryStream memoryStream = new MemoryStream();
            Package.SaveAs(memoryStream);
            memoryStream.Position = 0;
            return memoryStream;
        }
    }
}
