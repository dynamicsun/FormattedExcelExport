using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;


namespace FormattedExcelExport {
	public class TableBuilderComplex<TModel> where TModel : class {
		private readonly CultureInfo _culture;
		protected readonly string TableName;
		protected readonly Dictionary<string, Func<TModel, string>> Columns = new Dictionary<string, Func<TModel, string>>();

		public TableBuilderComplex(string tableName, CultureInfo culture) {
			TableName = tableName;
			_culture = culture;
		}

		public void RegisterColumnIf(bool condition, string columnName, Func<TModel, string> expression) {
			if (!condition)
				return;

			Columns.Add(columnName, expression);
		}

		public void RegisterColumnIf(bool condition, string columnName, Func<TModel, int?> expression) {
			RegisterColumnIf(condition, columnName, x => {
				int? value = expression(x);
				if (value == null)
					return string.Empty;

				return value.ToString();
			});
		}

		public void RegisterColumnIf(bool condition, string columnName, Func<TModel, decimal?> expression) {
			RegisterColumnIf(condition, columnName, x => {
				decimal? value = expression(x);
				if (value == null)
					return string.Empty;

				return string.Format(_culture, "{0:0.00}", value.Value);
			});
		}

		public void RegisterColumnIf(bool condition, string columnName, Func<TModel, DateTime?> expression) {
			RegisterColumnIf(condition, columnName, x => {
				DateTime? value = expression(x);
				if (value == null)
					return string.Empty;

				return value.Value.ToString(_culture.DateTimeFormat.ShortDatePattern, _culture);
			});
		}

		public void RegisterColumnIf(bool condition, string columnName, Func<TModel, bool?> expression) {
			RegisterColumnIf(condition, columnName, x => {
				bool? value = expression(x);
				if (value == null)
					return string.Empty;

				return value.Value ? "Да" : "Нет";
			});
		}

	}

	public sealed class CsvTableBuilderComplex<TModel> : TableBuilderComplex<TModel> where TModel : class {
		private readonly string _delimeter;
		public CsvTableBuilderComplex(string tableName, CultureInfo culture, string delimeter = "\t")
			: base(tableName, culture) {
			_delimeter = delimeter;
		}

		public void WriteParent(StringBuilder sb, TModel parent) {
			string[] header = Columns.Keys.ToArray();
			sb.Append(TableName).Append(_delimeter);

			int len = header.Length - 1;
			int i = 0;
			foreach (string s in header) {
				if (i < len)
					sb.Append(_delimeter);
				i++;
			}
			sb.AppendLine();

			sb.Append(_delimeter);
			i = 0;
			foreach (KeyValuePair<string, Func<TModel, string>> keyValuePair in Columns) {
				string s = keyValuePair.Value(parent);
				sb.Append(s);

				if (i < len)
					sb.Append(_delimeter);
				i++;
			}
			sb.AppendLine();
		}

		public void WriteChild(StringBuilder sb, IEnumerable<TModel> children) {
			string[] header = Columns.Keys.ToArray();
			sb.Append(TableName).Append(_delimeter);

			int len = header.Length - 1;
			int i = 0;
			foreach (string s in header) {
				sb.Append(s);

				if (i < len)
					sb.Append(_delimeter);
				i++;
			}
			sb.AppendLine();

			foreach (TModel child in children) {
				i = 0;
				sb.Append(_delimeter);
				foreach (KeyValuePair<string, Func<TModel, string>> keyValuePair in Columns) {
					string s = keyValuePair.Value(child);
					sb.Append(s);

					if (i < len)
						sb.Append(_delimeter);
					i++;
				}
				sb.AppendLine();
			}
		}
	}

	public sealed class XlsTableBuilderComplex<TModel> : TableBuilderComplex<TModel> where TModel : class {
		public XlsTableBuilderComplex(string tableName, CultureInfo culture)
			: base(tableName, culture) {
		}

		public void WriteParent(ref int rowIndex, TModel parent, HSSFWorkbook workbook, ISheet sheet) {
			ICellStyle cstyle = workbook.CreateCellStyle();
			cstyle.FillForegroundColor = HSSFColor.ROYAL_BLUE.index;
			cstyle.FillPattern = FillPatternType.SOLID_FOREGROUND;

			IFont font = workbook.CreateFont();
			font.Color = HSSFColor.WHITE.index;
			font.Boldweight = (short)FontBoldWeight.BOLD;
			cstyle.SetFont(font);
			cstyle.VerticalAlignment = VerticalAlignment.CENTER;

			IRow row = sheet.CreateRow(rowIndex);
			row.Height = 400;
			string[] header = Columns.Keys.ToArray();

			int colIndex = 0;
			var cl1 = row.CreateCell(colIndex);
			cl1.SetCellValue(TableName);
			cl1.CellStyle = cstyle;
			colIndex++;

			foreach (var s in header) {
				var cl = row.CreateCell(colIndex);
				cl.SetCellValue(s);
				sheet.SetColumnWidth(colIndex, 5000);
				cl.CellStyle = cstyle;
				colIndex++;
			}
			rowIndex++;

			cstyle = workbook.CreateCellStyle();
			font = workbook.CreateFont();
			font.Boldweight = (short)FontBoldWeight.BOLD;
			cstyle.SetFont(font);
			cstyle.VerticalAlignment = VerticalAlignment.CENTER;

			colIndex = 0;
			row = sheet.CreateRow(rowIndex);
			cl1 = row.CreateCell(colIndex);
			cl1.SetCellValue("");
			cl1.CellStyle = cstyle;
			colIndex++;


			foreach (KeyValuePair<string, Func<TModel, string>> keyValuePair in Columns) {
				var cl = row.CreateCell(colIndex);
				cl.SetCellValue(keyValuePair.Value(parent));
				cl.CellStyle = cstyle;
				colIndex++;
			}
			rowIndex++;
		}

		public void WriteChild(ref int rowIndex, IEnumerable<TModel> children, HSSFWorkbook workbook, ISheet sheet, short color) {
			ICellStyle cstyle = workbook.CreateCellStyle();
			cstyle.FillForegroundColor = color;
			cstyle.FillPattern = FillPatternType.SOLID_FOREGROUND;

			IRow row = sheet.CreateRow(rowIndex);
			string[] header = Columns.Keys.ToArray();

			int colIndex = 0;
			var cl = row.CreateCell(colIndex);
			cl.SetCellValue(TableName);
			cl.CellStyle = cstyle;

			colIndex++;
			foreach (var s in header) {
				row.CreateCell(colIndex).SetCellValue(s);
				colIndex++;
			}
			rowIndex++;

			foreach (var child in children) {
				colIndex = 0;
				row = sheet.CreateRow(rowIndex);
				row.CreateCell(colIndex).SetCellValue("");
				colIndex++;
				foreach (KeyValuePair<string, Func<TModel, string>> keyValuePair in Columns) {
					row.CreateCell(colIndex).SetCellValue(keyValuePair.Value(child));
					colIndex++;
				}
				rowIndex++;
			}
		}
	}
}
