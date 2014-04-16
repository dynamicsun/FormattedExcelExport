﻿using System.Collections.Generic;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;


namespace FormattedExcelExport {
	public class StyleSettings {
		public string FontName { get; set; }
		public short FontHeightInPoints { get; set; }
		public short BoldWeight { get; set; }
		public bool IsItalic { get; set; }
		public bool Underline { get; set; }
		public Color FontColor { get; set; }
		public Color BackgroundColor { get; set; }
		List<KeyValuePair<string, Color>> ColorsCollection { get; set; }

		public StyleSettings() {
			FontName = "Times New Roman";
			FontHeightInPoints = 7;
			BoldWeight = (short)FontBoldWeight.Normal;
			IsItalic = false;
			Underline = false;
			FontColor = new Color(255, 255, 255);
			BackgroundColor = new Color(0, 0, 0);
			ColorsCollection = new List<KeyValuePair<string, Color>> {
				new KeyValuePair<string, Color>("LIGHT_ORANGE", new Color(255, 217, 102)),
				new KeyValuePair<string, Color>("SEA_GREEN", new Color(198, 224, 180)),
				new KeyValuePair<string, Color>("VIOLET", new Color(136, 139, 252)),
				new KeyValuePair<string, Color>("BROWN", new Color(255, 124, 59)),
				new KeyValuePair<string, Color>("GREY_40_PERCENT", new Color(174, 170, 170))
			};
		}

		public enum FontBoldWeight : short {
			None = (short)0,
			Normal = (short)400,
			Bold = (short)700,
		}
		public class Color {
			public Color(byte red = 0, byte green = 0, byte blue = 0) {
				Red = red;
				Green = green;
				Blue = blue;
			}
			public byte Red { get; set; }
			public byte Green { get; set; }
			public byte Blue { get; set; }
		}
	}

	public class TableWriterStyle {
		public TableWriterStyle() {
			FontFactor = 300;
			FontAbsoluteTerm = 500;
			MaxColumnWidth = 25500;
			HeaderHeight = 400;

			HeaderCell = new StyleSettings();
		}

		public StyleSettings HeaderCell { get; set; }
		public StyleSettings RegularCell { get; set; }
		public StyleSettings HeaderChildCell { get; set; }
		public StyleSettings RegularChildCell { get; set; }

		public int FontFactor { get; set; }
		public int FontAbsoluteTerm { get; set; }
		public int MaxColumnWidth { get; set; }
		public short HeaderHeight { get; set; }
	}


	public interface ITableWriterStyle {
		short HeaderHeight { get; }
		ICellStyle HeaderCellStyle { get; }
		ICellStyle ChildHeaderCellStylePremier { get; }
		ICellStyle ChildHeaderCellStyleNext { get; }
		int MaxColumnWidth { get; }
		int FontFactor { get; }
		int FontAbsoluteTerm { get; }

		HSSFWorkbook Workbook { get; }
	}

	public class TableWriterStyleDefault : ITableWriterStyle {
		private readonly HSSFWorkbook _workbook;
		private readonly ICellStyle _headerCellStyle;
		private byte _colorIndex;
		private readonly List<short> _colors;
		public TableWriterStyleDefault() {
			_workbook = new HSSFWorkbook();

			var xx = new TableWriterStyle();

			IFont headerFont = _workbook.CreateFont();
			headerFont.Color = HSSFColor.WHITE.index;
			headerFont.Boldweight = (short)FontBoldWeight.BOLD;

			_headerCellStyle = _workbook.CreateCellStyle();
			_headerCellStyle.SetFont(headerFont);
			_headerCellStyle.FillForegroundColor = HSSFColor.ROYAL_BLUE.index;
			_headerCellStyle.FillPattern = FillPatternType.SOLID_FOREGROUND;
			_headerCellStyle.VerticalAlignment = VerticalAlignment.CENTER;

			HSSFPalette palette = _workbook.GetCustomPalette();
			palette.SetColorAtIndex(HSSFColor.LIGHT_ORANGE.index, 255, 217, 102);
			palette.SetColorAtIndex(HSSFColor.SEA_GREEN.index, 198, 224, 180);
			palette.SetColorAtIndex(HSSFColor.VIOLET.index, 136, 139, 252);
			palette.SetColorAtIndex(HSSFColor.BROWN.index, 255, 124, 59);
			palette.SetColorAtIndex(HSSFColor.GREY_40_PERCENT.index, 174, 170, 170);

			_colors = new List<short> {
				HSSFColor.LIGHT_ORANGE.index,
				HSSFColor.SEA_GREEN.index,
				HSSFColor.VIOLET.index,
				HSSFColor.BROWN.index,
				HSSFColor.GREY_40_PERCENT.index
			};
		}
		public short HeaderHeight {
			get { return 400; }
		}
		public ICellStyle HeaderCellStyle {
			get {
				return _headerCellStyle;
			}
		}
		public ICellStyle ChildHeaderCellStylePremier {
			get {
				ICellStyle cellStyle = _workbook.CreateCellStyle();
				cellStyle.FillPattern = FillPatternType.SOLID_FOREGROUND;
				cellStyle.FillForegroundColor = _colors.ElementAt(0);
				_colorIndex = 1;
				return cellStyle;
			}
		}
		public ICellStyle ChildHeaderCellStyleNext {
			get {
				ICellStyle cellStyle = _workbook.CreateCellStyle();
				cellStyle.FillPattern = FillPatternType.SOLID_FOREGROUND;
				cellStyle.FillForegroundColor = _colors.ElementAt(_colorIndex++);
				if (_colorIndex >= _colors.Count)
					_colorIndex = 0;
				return cellStyle;
			}
		}
		public int MaxColumnWidth {
			get { return 25500; }
		}
		public int FontFactor { get { return 300; } }
		public int FontAbsoluteTerm { get { return 500; } }
		public HSSFWorkbook Workbook {
			get { return _workbook; }
		}
	}
}
