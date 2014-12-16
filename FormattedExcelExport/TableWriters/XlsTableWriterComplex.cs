using System.Collections.Generic;
using System.Linq;
using FormattedExcelExport.Style;
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
	        for (var i = 0; i < Style.ColorsCollection.Count; i++) {
	            var color = Style.ColorsCollection.ElementAt(i);
	            if (color != null) {
	                var palette = Workbook.GetCustomPalette();
	                var similarColor = palette.FindSimilarColor(color.Red, color.Green, color.Blue);
                    var childHeaderCellStyle = ConvertToNpoiStyle(Style.HeaderChildCell);
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
	                var lastChildHeaderBuffer = _lastChildHeader;
	                WriteHeader(_lastParentHeader);
	                _lastChildHeader = lastChildHeaderBuffer;
	                if (_lastChildHeader != null) {
	                    WriteChildHeader(_lastChildHeader);
	                }
	            }
	        }
	    }

	    public void WriteHeader(params string[] cells) {
			var row = WorkSheet.CreateRow(RowIndex);
			HeaderCellStyle.VerticalAlignment = VerticalAlignment.Top;

			var columnIndex = 0;
			foreach (var cell in cells) {
				var newCell = row.CreateCell(columnIndex);
				newCell.SetCellValue(cell);
				newCell.CellStyle = HeaderCellStyle;
				columnIndex++;
			}
	        _lastParentHeader = cells;
	        _lastChildHeader = null;
			RowIndex++;
			_colorIndex = 0;
		}
        public void WriteRow(List<KeyValuePair<dynamic, TableWriterStyle>> cells, bool prependDelimeter = false) {
			WriteRowBase(cells, RowIndex, prependDelimeter);
			RowIndex++;
		}
		public void WriteChildHeader(params string[] cells) {
			var row = WorkSheet.CreateRow(RowIndex);
			var columnIndex = 0;
			var cellsList = cells.ToList();
			if (_colorIndex >= Style.ColorsCollection.Count)
				_colorIndex = 0;
            var childHeaderCellStyle = _childHeaderCellStyleList[_colorIndex];
			foreach (var cell in cellsList) {
				var newCell = row.CreateCell(columnIndex);
				newCell.SetCellValue(cell);
			    newCell.CellStyle = childHeaderCellStyle;
				columnIndex++;
			}
            _colorIndex++;
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