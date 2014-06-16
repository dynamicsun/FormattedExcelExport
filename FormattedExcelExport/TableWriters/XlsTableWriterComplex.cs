using System.Collections.Generic;
using System.Linq;
using FormattedExcelExport.Style;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;


namespace FormattedExcelExport.TableWriters {
	public sealed class XlsTableWriterComplex : XlsTableWriterBase, ITableWriterComplex {
	    private int _rowIndex;
	    private string[] _lastParentHeader;
	    private string[] _lastChildHeader;
		private byte _colorIndex;
	    private List<ICellStyle> _childHeaderCellStyleList;

	    public XlsTableWriterComplex(TableWriterStyle style) : base(style) {
            _childHeaderCellStyleList = new List<ICellStyle>();
	        for (int i = 0; i < Style.ColorsCollection.Count; i++) {
	            AdHocCellStyle.Color color = Style.ColorsCollection.ElementAt(i);
	            if (color != null) {
	                HSSFPalette palette = Workbook.GetCustomPalette();
	                HSSFColor similarColor = palette.FindSimilarColor(color.Red, color.Green, color.Blue);
                    ICellStyle childHeaderCellStyle = ConvertToNpoiStyle(Style.HeaderChildCell);
	                childHeaderCellStyle.FillForegroundColor = similarColor.GetIndex();
	                childHeaderCellStyle.FillPattern = FillPattern.SolidForeground;
                    _childHeaderCellStyleList.Add(childHeaderCellStyle);
	            }
	        }
	    }

	    public int RowIndex {
	        get { return _rowIndex; }
	        set {
	            if (_rowIndex < MaxRowIndex) {
	                _rowIndex = value;
	            }
	            else {
	                WorkSheet = Workbook.CreateSheet();
                    _rowIndex = 0;
	                string[] lastChildHeaderBuffer = _lastChildHeader;
	                WriteHeader(_lastParentHeader);
	                _lastChildHeader = lastChildHeaderBuffer;
	                if (_lastChildHeader != null) {
	                    WriteChildHeader(_lastChildHeader);
	                }
	            }
	        }
	    }

	    public void WriteHeader(params string[] cells) {
			IRow row = WorkSheet.CreateRow(RowIndex);
			row.Height = Style.HeaderHeight;
			
			HeaderCellStyle.VerticalAlignment = VerticalAlignment.Top;

			int columnIndex = 0;
			foreach (string cell in cells) {
				ICell newCell = row.CreateCell(columnIndex);
				newCell.SetCellValue(cell);
				newCell.CellStyle = HeaderCellStyle;
				columnIndex++;
			}
	        _lastParentHeader = cells;
	        _lastChildHeader = null;
			RowIndex++;
			_colorIndex = 0;
		}
		public void WriteRow(IEnumerable<KeyValuePair<string, TableWriterStyle>> cells, bool prependDelimeter = false) {
			IRow row = WorkSheet.CreateRow(RowIndex);
			int columnIndex = 0;
			if (prependDelimeter) {
				ICell newCell = row.CreateCell(columnIndex);
				newCell.SetCellValue("");
				newCell.CellStyle = CellStyle;

				columnIndex++;
			}
			foreach (KeyValuePair<string, TableWriterStyle> cell in cells) {
				ICell newCell = row.CreateCell(columnIndex);
				newCell.SetCellValue(cell.Key);

				if (cell.Value != null) {
					ICellStyle customCellStyle = ConvertToNpoiStyle(cell.Value.RegularCell);
					newCell.CellStyle = customCellStyle;
				}
				else {
					newCell.CellStyle = CellStyle;
				}
				columnIndex++;
			}
			RowIndex++;
		}
		public void WriteChildHeader(params string[] cells) {
			IRow row = WorkSheet.CreateRow(RowIndex);
			int columnIndex = 0;
			List<string> cellsList = cells.ToList();
			//_childHeaderCellStyle = ConvertToNpoiStyle(Style.HeaderChildCell);

			if (_colorIndex >= Style.ColorsCollection.Count)
				_colorIndex = 0;

			/*AdHocCellStyle.Color color = Style.ColorsCollection.ElementAt(_colorIndex);
			if (color != null) {
				HSSFPalette palette = Workbook.GetCustomPalette();
				HSSFColor similarColor = palette.FindSimilarColor(color.Red, color.Green, color.Blue);
				_childHeaderCellStyle.FillForegroundColor = similarColor.GetIndex();
				_childHeaderCellStyle.FillPattern = FillPattern.SolidForeground;
				_colorIndex++;
			}*/
            ICellStyle childHeaderCellStyle = _childHeaderCellStyleList[_colorIndex];
			foreach (string cell in cellsList) {
				ICell newCell = row.CreateCell(columnIndex);
				newCell.SetCellValue(cell);
			    newCell.CellStyle = childHeaderCellStyle;
				columnIndex++;
			}
            _colorIndex++;
		    _lastChildHeader = cells;
			RowIndex++;
		}
		public void WriteChildRow(IEnumerable<KeyValuePair<string, TableWriterStyle>> cells, bool prependDelimeter = false) {
			IRow row = WorkSheet.CreateRow(RowIndex);

			int columnIndex = 0;
			if (prependDelimeter) {
				ICell newCell = row.CreateCell(columnIndex);
				newCell.SetCellValue("");
				newCell.CellStyle = CellStyle;

				columnIndex++;
			}
			foreach (KeyValuePair<string, TableWriterStyle> cell in cells) {
				ICell newCell = row.CreateCell(columnIndex);
				newCell.SetCellValue(cell.Key);

				if (cell.Value != null) {
					ICellStyle customCellStyle = ConvertToNpoiStyle(cell.Value.RegularChildCell);
					newCell.CellStyle = customCellStyle;
				}
				else {
					newCell.CellStyle = CellStyle;
				}
				columnIndex++;
			}
			RowIndex++;
		}
	}
}