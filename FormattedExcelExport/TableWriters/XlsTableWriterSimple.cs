using System.Collections.Generic;
using FormattedExcelExport.Style;
using NPOI.SS.UserModel;


namespace FormattedExcelExport.TableWriters {
	public sealed class XlsTableWriterSimple : XlsTableWriterBase, ITableWriterSimple {
        public XlsTableWriterSimple(TableWriterStyle style = null) : base(style) {}
        private int _rowIndex;
        private List<string> _lastHeader;

	    public int RowIndex {
	        get { return _rowIndex; }
	        set {
	            if (_rowIndex < MaxRowIndex) {
	                _rowIndex = value;
	            }
	            else {
	                WorkSheet = Workbook.CreateSheet();
                    _rowIndex = 0;
	                WriteHeader(_lastHeader);	                
	            }
	        }
	    }

	    public void WriteHeader(List<string> cells) {
			var row = WorkSheet.CreateRow(RowIndex);
			CellStyle.VerticalAlignment = VerticalAlignment.Top;

			var columnIndex = 0;
			foreach (var cell in cells) {
				var newCell = row.CreateCell(columnIndex);
				newCell.SetCellValue(cell);
				newCell.CellStyle = HeaderCellStyle;
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