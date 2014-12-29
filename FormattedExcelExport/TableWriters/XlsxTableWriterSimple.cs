using System.Collections.Generic;
using FormattedExcelExport.Style;

namespace FormattedExcelExport.TableWriters {
    public sealed class XlsxTableWriterSimple : XlsxTableWriterBase, ITableWriterSimple {
        public XlsxTableWriterSimple(TableWriterStyle style) : base(style) {}
        private int _rowIndex = 1;
        public int RowIndex {
            get { return _rowIndex; }
            set {
                if (_rowIndex < MaxRowIndex) {
                    _rowIndex = value;
                } else {
                    SheetNumber = SheetNumber + 1;
                    WorkSheet = Package.Workbook.Worksheets.Add("Sheet " + SheetNumber);
                    _rowIndex = 1;
                    WriteHeader(LastHeader);
                }
            }
        }
        public void WriteHeader(IEnumerable<string> cells) {
            WriteHeaderBase(cells, RowIndex);
            RowIndex++;
        }

        public void WriteRow(List<KeyValuePair<dynamic, TableWriterStyle>> cells) {
            WriteRowBase(cells, RowIndex);
            RowIndex++;
        }
    }
}
