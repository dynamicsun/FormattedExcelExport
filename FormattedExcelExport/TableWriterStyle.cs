using System.Collections.Generic;

namespace FormattedExcelExport {
	public class StyleSettings {
		public string FontName { get; set; }
		public short FontHeightInPoints { get; set; }
		public FontBoldWeight BoldWeight { get; set; }
		public bool Italic { get; set; }
		public bool Underline { get; set; }
		public Color FontColor { get; set; }
		public Color BackgroundColor { get; set; }

		public StyleSettings() {
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

			HeaderCell = new StyleSettings {
				BoldWeight = StyleSettings.FontBoldWeight.Bold,
				FontColor = new StyleSettings.Color(255, 255, 255),
				BackgroundColor = new StyleSettings.Color(2, 101, 203)
			};

			RegularCell = new StyleSettings {
				FontColor = new StyleSettings.Color()
			};
			HeaderChildCell = new StyleSettings();
			RegularChildCell = new StyleSettings();

			ColorsCollection = new List<StyleSettings.Color> {
				new StyleSettings.Color(255, 217, 102),
				new StyleSettings.Color(198, 224, 180),
				new StyleSettings.Color(136, 139, 252),
				new StyleSettings.Color(255, 124, 59),
				new StyleSettings.Color(174, 170, 170),
			};
		}
		public List<StyleSettings.Color> ColorsCollection { get; set; }
		public StyleSettings HeaderCell { get; set; }
		public StyleSettings RegularCell { get; set; }
		public StyleSettings HeaderChildCell { get; set; }
		public StyleSettings RegularChildCell { get; set; }

		public int FontFactor { get; set; }
		public int FontAbsoluteTerm { get; set; }
		public int MaxColumnWidth { get; set; }
		public short HeaderHeight { get; set; }
	}
}
