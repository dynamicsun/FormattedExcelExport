using System.Collections.Generic;
using System.IO;
using System.Linq;
using FormattedExcelExport.Style;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;


namespace FormattedExcelExport.TableWriters {
	public sealed class XlsxTableWriterComplex : XlsxTableWriterBase, ITableWriterComplex {
        private int _rowIndex;
        private string[] _lastParentHeader;
        private string[] _lastChildHeader;
        private byte _colorIndex;
        private List<ICellStyle> _childHeaderCellStyleList;

	    public XlsxTableWriterComplex(TableWriterStyle style)
	        : base(style) {
	        /*_childHeaderCellStyleList = new List<ICellStyle>();
            for (int i = 0; i < Style.ColorsCollection.Count; i++) {
                AdHocCellStyle.Color color = Style.ColorsCollection.ElementAt(i);
                if (color != null) {
                    
                    XSSFCellStyle childHeaderCellStyle = (XSSFCellStyle) ConvertToNpoiStyle(Style.HeaderChildCell);
                    childHeaderCellStyle.SetFillForegroundColor(new XSSFColor(new[] { color.Red, color.Green, color.Blue }));
                    childHeaderCellStyle.FillPattern = FillPattern.SolidForeground;
                    _childHeaderCellStyleList.Add(childHeaderCellStyle);
                }
            }*/
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
                    WriteHeader(_lastParentHeader);
                    WriteChildHeader(_lastChildHeader);                    
	            }
	        }
	    }

		
		
		public void WriteHeader(params string[] cells) {
			IRow row = WorkSheet.CreateRow(RowIndex);
			row.Height = Style.HeaderHeight;

			ICellStyle cellStyle = ConvertToNpoiStyle(Style.HeaderCell);
			cellStyle.FillPattern = FillPattern.SolidForeground;
			cellStyle.VerticalAlignment = VerticalAlignment.Center;

			int columnIndex = 0;
			foreach (string cell in cells) {
				ICell newCell = row.CreateCell(columnIndex);
				newCell.SetCellValue(cell);
				newCell.CellStyle = cellStyle;

				XSSFCellStyle cs = (XSSFCellStyle)newCell.CellStyle;
				cs.SetFillForegroundColor(new XSSFColor(new[] { Style.HeaderCell.BackgroundColor.Red, Style.HeaderCell.BackgroundColor.Green, Style.HeaderCell.BackgroundColor.Blue }));

				columnIndex++;
			}
		    _lastParentHeader = cells;
			RowIndex++;
			_colorIndex = 0;
		}
		public void WriteRow(IEnumerable<KeyValuePair<string, TableWriterStyle>> cells, bool prependDelimeter = false) {
			IRow row = WorkSheet.CreateRow(RowIndex);
			ICellStyle cellStyle = ConvertToNpoiStyle(Style.RegularCell);

			int columnIndex = 0;
			if (prependDelimeter) {
				ICell newCell = row.CreateCell(columnIndex);
				newCell.SetCellValue("");
				newCell.CellStyle = cellStyle;

				if (Style.RegularCell.BackgroundColor != null) {
					XSSFCellStyle cs = (XSSFCellStyle)newCell.CellStyle;
					cs.SetFillForegroundColor(new XSSFColor(new[] { Style.RegularCell.BackgroundColor.Red, Style.RegularCell.BackgroundColor.Green, Style.RegularCell.BackgroundColor.Blue }));
					cellStyle.FillPattern = FillPattern.SolidForeground;
				}
				columnIndex++;
			}
			foreach (KeyValuePair<string, TableWriterStyle> cell in cells) {
				ICell newCell = row.CreateCell(columnIndex);
				newCell.SetCellValue(cell.Key);

				if (cell.Value != null) {
					ICellStyle customCellStyle = ConvertToNpoiStyle(cell.Value.RegularCell);
					newCell.CellStyle = customCellStyle;
					if (cell.Value.RegularCell.BackgroundColor != null) {
						XSSFCellStyle cs = (XSSFCellStyle)newCell.CellStyle;
						cs.SetFillForegroundColor(new XSSFColor(new[] { cell.Value.RegularCell.BackgroundColor.Red, cell.Value.RegularCell.BackgroundColor.Green, cell.Value.RegularCell.BackgroundColor.Blue }));
						cs.FillPattern = FillPattern.SolidForeground;
					}
				}
				else {
					newCell.CellStyle = cellStyle;
				}
				columnIndex++;
			}
			RowIndex++;
		}
		public void WriteChildHeader(params string[] cells) {
			IRow row = WorkSheet.CreateRow(RowIndex);
			int columnIndex = 0;
			List<string> cellsList = cells.ToList();
			ICellStyle cellStyle = ConvertToNpoiStyle(Style.HeaderChildCell);

			if (_colorIndex >= Style.ColorsCollection.Count)
				_colorIndex = 0;

			AdHocCellStyle.Color color = Style.ColorsCollection.ElementAt(_colorIndex);
			_colorIndex++;
            //ICellStyle childHeaderCellStyle = _childHeaderCellStyleList[_colorIndex];
			foreach (string cell in cellsList) {
				ICell newCell = row.CreateCell(columnIndex);
				newCell.SetCellValue(cell);
				newCell.CellStyle = cellStyle;
			    //newCell.CellStyle = childHeaderCellStyle;
				if (color != null) {
					XSSFCellStyle cs = (XSSFCellStyle)newCell.CellStyle;
					cs.SetFillForegroundColor(new XSSFColor(new[] { color.Red, color.Green, color.Blue }));
					cs.FillPattern = FillPattern.SolidForeground;
				}

				columnIndex++;
			}
		    //_colorIndex++;
		    _lastChildHeader = cells;
			RowIndex++;
		}
		public void WriteChildRow(IEnumerable<KeyValuePair<string, TableWriterStyle>> cells, bool prependDelimeter = false) {
			IRow row = WorkSheet.CreateRow(RowIndex);

			ICellStyle cellStyle = ConvertToNpoiStyle(Style.RegularChildCell);

			int columnIndex = 0;
			if (prependDelimeter) {
				ICell newCell = row.CreateCell(columnIndex);
				newCell.SetCellValue("");
				newCell.CellStyle = cellStyle;

				columnIndex++;
			}

			foreach (KeyValuePair<string, TableWriterStyle> cell in cells) {
				ICell newCell = row.CreateCell(columnIndex);
				newCell.SetCellValue(cell.Key);

				if (cell.Value != null) {
					ICellStyle customCellStyle = ConvertToNpoiStyle(cell.Value.RegularChildCell);
					newCell.CellStyle = customCellStyle;

					if (cell.Value.RegularChildCell.BackgroundColor != null) {
						XSSFCellStyle cs = (XSSFCellStyle)newCell.CellStyle;
						cs.SetFillForegroundColor(new XSSFColor(new[] { cell.Value.RegularChildCell.BackgroundColor.Red, cell.Value.RegularChildCell.BackgroundColor.Green, cell.Value.RegularChildCell.BackgroundColor.Blue }));
						cs.FillPattern = FillPattern.SolidForeground;
					}
				}
				else {
					newCell.CellStyle = cellStyle;
				}
				columnIndex++;
			}
			RowIndex++;
		}

		public MemoryStream GetStream() {
			MemoryStream memoryStream = new MemoryStream();
//			Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

			FileStream sw = File.Create("TestComplex.xlsx");
			Workbook.Write(sw);


			//			Workbook.Write(memoryStream);
			//			memoryStream.Position = 0;
			return memoryStream;
		}
	}
}
