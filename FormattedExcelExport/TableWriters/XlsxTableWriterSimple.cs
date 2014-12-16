using System.Collections.Generic;
using System.Drawing;
using FormattedExcelExport.Style;
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
            _lastHeader = cells;
            RowIndex++;
        }

        public void WriteRow(List<KeyValuePair<dynamic, TableWriterStyle>> cells) {
            WriteRowBase(cells, RowIndex);
            RowIndex++;
        }
    }
}
