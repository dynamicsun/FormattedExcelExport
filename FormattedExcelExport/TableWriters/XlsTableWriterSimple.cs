using System.Collections.Generic;
using FormattedExcelExport.Style;


namespace FormattedExcelExport.TableWriters {
	public sealed class XlsTableWriterSimple : XlsTableWriterBase, ITableWriterSimple {
        public XlsTableWriterSimple(TableWriterStyle style = null) : base(style) {}
        private int _rowIndex;

	    public int RowIndex {
	        get { return _rowIndex; }
	        set {
	            if (_rowIndex < MaxRowIndex) {
	                _rowIndex = value;
	            }
	            else {
	                WorkSheet = Workbook.CreateSheet();
                    _rowIndex = 0;
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