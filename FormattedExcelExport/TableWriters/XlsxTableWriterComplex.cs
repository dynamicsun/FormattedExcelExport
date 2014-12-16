using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using FormattedExcelExport.Style;
using OfficeOpenXml.Style;

namespace FormattedExcelExport.TableWriters {
    public sealed class XlsxTableWriterComplex : XlsxTableWriterBase, ITableWriterComplex {
        private int _rowIndex = 1;
        private string[] _lastParentHeader;
        private string[] _lastChildHeader;
        private byte _colorIndex;

        public int RowIndex {
            get { return _rowIndex; }
            set {
                if (_rowIndex < MaxRowIndex) {
                    _rowIndex = value;
                } else {
                    SheetNumber = SheetNumber + 1;
                    WorkSheet = Package.Workbook.Worksheets.Add("Sheet " + SheetNumber);
                    _rowIndex = 1;
                    WriteHeader(_lastParentHeader);
                    WriteChildHeader(_lastChildHeader);
                }
            }
        }

        public XlsxTableWriterComplex(TableWriterStyle style) : base(style) {
        }

        public void WriteHeader(params string[] cells) {
            var font = ConvertCellStyle(Style.HeaderCell);
            var columnIndex = 1;
            foreach (var cell in cells) {
                var newCell = WorkSheet.Cells[RowIndex, columnIndex];
                newCell.Value = cell;
                newCell.Style.Font.SetFromFont(font);
                newCell.Style.Font.Color.SetColor(Color.FromArgb(Style.HeaderCell.FontColor.Red, Style.HeaderCell.FontColor.Green, Style.HeaderCell.FontColor.Blue));
                newCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                newCell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(Style.HeaderCell.BackgroundColor.Red, Style.HeaderCell.BackgroundColor.Green, Style.HeaderCell.BackgroundColor.Blue));
                newCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                columnIndex++;
            }
            _lastParentHeader = cells;
            RowIndex++;
            _colorIndex = 0;
        }

        public void WriteRow(List<KeyValuePair<dynamic, TableWriterStyle>> cells, bool prependDelimeter = false) {
            WriteRowBase(cells, RowIndex, prependDelimeter);
            RowIndex++;
        }

        public void WriteChildHeader(params string[] cells) {
            var font = ConvertCellStyle(Style.HeaderChildCell);
            var columnIndex = 1;

            if (_colorIndex >= Style.ColorsCollection.Count) _colorIndex = 0;
            var color = Style.ColorsCollection.ElementAt(_colorIndex);
            _colorIndex++;

            foreach (var cell in cells) {
                var newCell = WorkSheet.Cells[RowIndex, columnIndex];
                newCell.Value = cell;
                newCell.Style.Font.SetFromFont(font);
                newCell.Style.Font.Color.SetColor(Color.FromArgb(Style.HeaderChildCell.FontColor.Red, Style.HeaderChildCell.FontColor.Green, Style.HeaderChildCell.FontColor.Blue));
                newCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                newCell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(color.Red, color.Green, color.Blue));
                newCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                columnIndex++;
            }
            _lastChildHeader = cells;
            RowIndex++;
        }

        public void WriteChildRow(IEnumerable<KeyValuePair<dynamic, TableWriterStyle>> cells, bool prependDelimeter = false) {
            const bool isChildRow = true;
            WriteRowBase(cells, RowIndex, prependDelimeter, isChildRow);
            RowIndex++;
        }
    }
}
