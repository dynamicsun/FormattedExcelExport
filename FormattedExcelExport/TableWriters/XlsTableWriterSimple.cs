using System;
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
			IRow row = WorkSheet.CreateRow(RowIndex);
			//row.Height = Style.HeaderHeight;
			CellStyle.VerticalAlignment = VerticalAlignment.Top;

			int columnIndex = 0;
			foreach (string cell in cells) {
				ICell newCell = row.CreateCell(columnIndex);
				newCell.SetCellValue(cell);
				newCell.CellStyle = HeaderCellStyle;
				columnIndex++;
			}
	        _lastHeader = cells;
			RowIndex++;
		}
        public void WriteRow(List<KeyValuePair<dynamic, TableWriterStyle>> cells) {
			IRow row = WorkSheet.CreateRow(RowIndex);
		    int columnIndex = 0;
            
            
            foreach (KeyValuePair<dynamic, TableWriterStyle> cell in cells) {
				ICell newCell = row.CreateCell(columnIndex);
				if (cell.Key != null)
					newCell.SetCellValue(cell.Key);
                if (cell.Key != null && cell.Key is DateTime?) {
                    newCell.CellStyle = DateCellStyle;
                }
                else {
                    if (cell.Value != null) {
                        ICellStyle customCellStyle = ConvertToNpoiStyle(cell.Value.RegularCell);
                        if (cell.Value.RegularCell.BackgroundColor != null) {
                            newCell.CellStyle = customCellStyle;
                        }
                        if (cell.Value.RegularChildCell.BackgroundColor != null) {
                            newCell.CellStyle = customCellStyle;
                        }
                    }
                    else {
                        newCell.CellStyle = CellStyle;
                    }
                }
                columnIndex++;
			}
			RowIndex++;
		}
	}
}