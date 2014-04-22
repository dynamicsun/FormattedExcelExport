using System.Collections.Generic;


namespace FormattedExcelExport.Style {
	public class AdHocCellStyle {
		public string FontName { get; set; }
		public short FontHeightInPoints { get; set; }
		public FontBoldWeight BoldWeight { get; set; }
		public bool Italic { get; set; }
		public bool Underline { get; set; }
		public Color FontColor { get; set; }
		public Color BackgroundColor { get; set; }

		public AdHocCellStyle() {
			FontName = "Arial";
			FontHeightInPoints = 10;
			BoldWeight = FontBoldWeight.Normal;
			Italic = false;
			Underline = false;
			FontColor = new Color();
		}

		public enum FontBoldWeight {
			None = 0,
			Normal = 400,
			Bold = 700,
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

			HeaderCell = new AdHocCellStyle {
				BoldWeight = AdHocCellStyle.FontBoldWeight.Bold,
				FontColor = new AdHocCellStyle.Color(255, 255, 255),
				BackgroundColor = new AdHocCellStyle.Color(2, 101, 203)
			};

			RegularCell = new AdHocCellStyle {
				FontColor = new AdHocCellStyle.Color()
			};
			HeaderChildCell = new AdHocCellStyle();
			RegularChildCell = new AdHocCellStyle();

			ColorsCollection = new List<AdHocCellStyle.Color> {
				new AdHocCellStyle.Color(255, 217, 102),
				new AdHocCellStyle.Color(198, 224, 180),
				new AdHocCellStyle.Color(136, 139, 252),
				new AdHocCellStyle.Color(255, 124, 59),
				new AdHocCellStyle.Color(174, 170, 170),
			};
		}
		public List<AdHocCellStyle.Color> ColorsCollection { get; set; }
		public AdHocCellStyle HeaderCell { get; set; }
		public AdHocCellStyle RegularCell { get; set; }
		public AdHocCellStyle HeaderChildCell { get; set; }
		public AdHocCellStyle RegularChildCell { get; set; }

		public int FontFactor { get; set; }
		public int FontAbsoluteTerm { get; set; }
		public int MaxColumnWidth { get; set; }
		public short HeaderHeight { get; set; }
	}
}
