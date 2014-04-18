using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;


namespace FormattedExcelExport {
	public class TableConfiguration {
		public string Title { get; set; }
		public readonly Dictionary<string, AggregatedContainer> ColumnsMap = new Dictionary<string, AggregatedContainer>();
		public readonly List<ChildTableConfiguration> ChildrenMap = new List<ChildTableConfiguration>();
	}

	public class ChildTableConfiguration : TableConfiguration {
		public Func<object, IEnumerable<object>> Getter { get; set; }
	}

	public class AggregatedContainer {
		public AggregatedContainer(Func<object, string> valueFunc, Func<object, bool> conditionFunc, TableWriterStyle style) {
			ValueFunc = valueFunc;
			ConditionFunc = conditionFunc;
			Style = style;
		}
		public Func<object, string> ValueFunc { get; set; }
		public TableWriterStyle Style { get; set; }
		public Func<object, bool> ConditionFunc { get; set; }
	}

	public sealed class TableConfigurationBuilder<TModel> {
		public class ConditionTheme {
			public ConditionTheme(TableWriterStyle style, Func<TModel, bool> condition) {
				Style = style;
				Condition = condition;
			}
			public Func<TModel, bool> Condition { get; set; }
			public TableWriterStyle Style { get; set; }
		}

		public TableConfiguration Value { get; set; }
		private readonly CultureInfo _culture;

		public TableConfigurationBuilder(string title, CultureInfo culture) {
			Value = new TableConfiguration { Title = title };
			_culture = culture;
		}
		public TableConfigurationBuilder(string title, Func<object, IEnumerable<object>> getter, CultureInfo culture) {
			Value = new ChildTableConfiguration { Getter = getter, Title = title };
			_culture = culture;
		}

		public void RegisterColumn(string header, Func<TModel, string> getter, ConditionTheme conditionTheme = null) {
			RegisterColumnIf(true, header, getter, conditionTheme);
		}

		public void RegisterColumn(string header, Func<TModel, int?> getter, ConditionTheme conditionTheme = null) {
			RegisterColumnIf(true, header, getter, conditionTheme);
		}

		public void RegisterColumn(string header, Func<TModel, decimal?> getter, ConditionTheme conditionTheme = null) {
			RegisterColumnIf(true, header, getter, conditionTheme);
		}

		public void RegisterColumn(string header, Func<TModel, DateTime?> getter, ConditionTheme conditionTheme = null) {
			RegisterColumnIf(true, header, getter, conditionTheme);
		}

		public void RegisterColumn(string header, Func<TModel, bool?> getter, ConditionTheme conditionTheme = null) {
			RegisterColumnIf(true, header, getter, conditionTheme);
		}

		public void RegisterColumnIf(bool expression, string header, Func<TModel, string> getter, ConditionTheme conditionTheme = null) {
			if (!expression)
				return;

			if(conditionTheme == null)
				conditionTheme = new ConditionTheme(null, x => false);

			Value.ColumnsMap.Add(header, new AggregatedContainer(x => getter((TModel)x), y => conditionTheme.Condition((TModel)y), conditionTheme.Style));
		}

		public void RegisterColumnIf(bool expression, string header, Func<TModel, int?> getter, ConditionTheme conditionTheme = null) {
			RegisterColumnIf(expression, header, x => getter(x).ToString(), conditionTheme);
		}

		public void RegisterColumnIf(bool expression, string header, Func<TModel, decimal?> getter, ConditionTheme conditionTheme = null) {
			RegisterColumnIf(expression, header, x => {
				decimal? value = getter(x);

				return value.HasValue ? string.Format(_culture, "{0:C}", value.Value) : string.Empty;
			}, conditionTheme);
		}

		public void RegisterColumnIf(bool expression, string header, Func<TModel, DateTime?> getter, ConditionTheme conditionTheme = null) {
			RegisterColumnIf(expression, header, x => {
				DateTime? value = getter(x);

				return value.HasValue ? value.Value.ToString(_culture.DateTimeFormat.LongDatePattern, _culture) : string.Empty;
			}, conditionTheme);
		}

		public void RegisterColumnIf(bool expression, string header, Func<TModel, bool?> getter, ConditionTheme conditionTheme = null) {
			RegisterColumnIf(expression, header, x => {
				bool? value = getter(x);
				if (!value.HasValue)
					return string.Empty;

				return value.Value ? "Да" : "Нет";
			}, conditionTheme);
		}

		public TableConfigurationBuilder<TChildModel> RegisterChild<TChildModel>(string title, Func<TModel, IEnumerable<TChildModel>> getter) {
			TableConfigurationBuilder<TChildModel> tableConfigurationBuilder = new TableConfigurationBuilder<TChildModel>(title, x => {
				if (getter((TModel)x) != null)
					return getter((TModel)x).Cast<object>();

				return new List<object>();
			}, _culture);
			Value.ChildrenMap.Add((ChildTableConfiguration)tableConfigurationBuilder.Value);

			return tableConfigurationBuilder;
		}
	}

	public abstract class ExcelTableWriterBase {
		protected int RowIndex;
		protected readonly HSSFWorkbook Workbook = new HSSFWorkbook();
		protected readonly ISheet WorkSheet;
		protected readonly TableWriterStyle Style;

		protected ExcelTableWriterBase(TableWriterStyle style) {
			Style = style;
			WorkSheet = Workbook.CreateSheet();
		}

		public void AutosizeColumns() {
			var columnLengths = new List<int>();

			for (int columnNum = 0; columnNum < WorkSheet.GetRow(0).LastCellNum; columnNum++) {
				int columnMaximumLength = 0;
				for (int rowNum = 0; rowNum <= WorkSheet.LastRowNum; rowNum++) {
					IRow currentRow = WorkSheet.GetRow(rowNum);

					if (!currentRow.Cells.Any()) continue;
					ICell cell = currentRow.GetCell(columnNum);
					if (cell == null) continue;

					if (cell.StringCellValue.Length > columnMaximumLength)
						columnMaximumLength = cell.StringCellValue.Length;
				}
				columnLengths.Add(columnMaximumLength);
			}

			for (int i = 0; i < WorkSheet.GetRow(0).LastCellNum; i++) {
				int width = columnLengths.ElementAt(i) * Style.FontFactor + Style.FontAbsoluteTerm;
				WorkSheet.SetColumnWidth(i, width < Style.MaxColumnWidth ? width : Style.MaxColumnWidth);
			}
		}

		protected ICellStyle ConvertToNpoiStyle(AdHocCellStyle adHocCellStyle) {
			IFont cellFont = Workbook.CreateFont();

			cellFont.FontName = adHocCellStyle.FontName;
			cellFont.FontHeightInPoints = adHocCellStyle.FontHeightInPoints;
			cellFont.IsItalic = adHocCellStyle.Italic;
			cellFont.Underline = adHocCellStyle.Underline ? FontUnderline.SINGLE.ByteValue : FontUnderline.NONE.ByteValue;
			cellFont.Boldweight = (short)adHocCellStyle.BoldWeight;

			HSSFPalette palette = Workbook.GetCustomPalette();
			HSSFColor similarColor = palette.FindSimilarColor(adHocCellStyle.FontColor.Red, adHocCellStyle.FontColor.Green, adHocCellStyle.FontColor.Blue);
			cellFont.Color = similarColor.GetIndex();

			ICellStyle cellStyle = Workbook.CreateCellStyle();
			cellStyle.SetFont(cellFont);

			if (adHocCellStyle.BackgroundColor != null) {
				similarColor = palette.FindSimilarColor(adHocCellStyle.BackgroundColor.Red, adHocCellStyle.BackgroundColor.Green, adHocCellStyle.BackgroundColor.Blue);
				cellStyle.FillForegroundColor = similarColor.GetIndex();
				cellStyle.FillPattern = FillPatternType.SOLID_FOREGROUND;
			}
			return cellStyle;
		}

		public MemoryStream GetStream() {
			MemoryStream memoryStream = new MemoryStream();
			Workbook.Write(memoryStream);
			memoryStream.Position = 0;
			return memoryStream;
		}
	}
}
