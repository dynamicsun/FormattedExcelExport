using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;


namespace FormattedExcelExport {
	public interface ITableWriter {
		void WriteHeader(params string[] cells);
		void WriteRow(params string[] cells);
		void AutosizeColumns();
		MemoryStream GetStream();
	}
	public sealed class CsvTableWriter : ITableWriter {
		private StringBuilder _sb = new StringBuilder();
		private readonly string _delimeter;
		public CsvTableWriter(string delimeter = "\t") {
			_delimeter = delimeter;
		}

		public void WriteHeader(params string[] cells) {
			WriteRow(cells);
		}
		public void WriteRow(params string[] cells) {
			int len = cells.Length - 1;
			int i = 0;
			foreach (string cell in cells) {
				_sb.Append(cell);

				if (i < len)
					_sb.Append(_delimeter);
				i++;
			}
			_sb.AppendLine();
		}
		public void AutosizeColumns() { }
		public MemoryStream GetStream() {
			MemoryStream ms = new MemoryStream();
			StreamWriter sw = new StreamWriter(ms, Encoding.UTF8);
			sw.WriteLine(_sb.ToString());
			sw.Flush();
			ms.Position = 0;
			return ms;
		}
	}

	public sealed class XlsTableWriter : ITableWriter {
		private int _rowIndex;
		readonly HSSFWorkbook _workbook = new HSSFWorkbook();
		ISheet _sheet;

		public XlsTableWriter() {
			_sheet = _workbook.CreateSheet();
		}

		public void WriteHeader(params string[] cells) {
			ICellStyle cstyle = _workbook.CreateCellStyle();
			cstyle.FillForegroundColor = HSSFColor.ROYAL_BLUE.index;
			cstyle.FillPattern = FillPatternType.SOLID_FOREGROUND;

			IFont font = _workbook.CreateFont();
			font.Color = HSSFColor.WHITE.index;
			font.Boldweight = (short)FontBoldWeight.BOLD;
			cstyle.SetFont(font);
			cstyle.VerticalAlignment = VerticalAlignment.CENTER;

			IRow row = _sheet.CreateRow(_rowIndex);
			row.Height = 400;

			int colIndex = 0;
			foreach (var cell in cells) {
				var cl = row.CreateCell(colIndex);
				cl.SetCellValue(cell);
				cl.CellStyle = cstyle;
				colIndex++;
			}
			_rowIndex++;
		}
		public void WriteRow(params string[] cells) {
			IRow row = _sheet.CreateRow(_rowIndex);
			int colIndex = 0;
			foreach (var cell in cells) {
				row.CreateCell(colIndex).SetCellValue(cell);
				colIndex++;
			}
			_rowIndex++;
		}

		public void AutosizeColumns() {
			const int maxExcelColumnWidth = 25500;
			var colLengths = new List<int>();

			for (int i = 0; i < _sheet.GetRow(0).LastCellNum; i++) {
				int colMaxLenght = 0;
				for (int j = 0; j < _sheet.LastRowNum; j++) {
					IRow row = _sheet.GetRow(j);
					if (row.Cells.Any()) {
						var cell = row.GetCell(i);

						if (cell != null) {
							if (cell.StringCellValue.Length > colMaxLenght)
								colMaxLenght = cell.StringCellValue.Length;
						}
					}
				}
				colLengths.Add(colMaxLenght);
			}


			for (int i = 0; i < _sheet.GetRow(0).LastCellNum; i++) {
				int width = colLengths.ElementAt(i) * 300 + 500;
				_sheet.SetColumnWidth(i, width < maxExcelColumnWidth ? width : maxExcelColumnWidth);
			}
		}

		public MemoryStream GetStream() {
			MemoryStream ms = new MemoryStream();
			_workbook.Write(ms);
			ms.Position = 0;
			return ms;
		}
	}

	public sealed class TableConfigurationBuilder<TModel> {
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

		public void RegisterColumn(string header, Func<TModel, string> getter) {
			Value.ColumnsMap.Add(header, x => getter((TModel)x));
		}

		public void RegisterColumn(string header, Func<TModel, decimal?> getter) {
			RegisterColumn(header, x => {
				decimal? value = getter(x);
				return value.HasValue ? string.Format(_culture, "{0:C}", value.Value) : string.Empty;
			});
		}

		public void RegisterColumnIf(bool expression, string header, Func<TModel, string> getter) {
			if (!expression)
				return;

			Value.ColumnsMap.Add(header, x => getter((TModel)x));
		}

		public void RegisterColumnIf(bool expression, string header, Func<TModel, int?> getter) {
			RegisterColumnIf(expression, header, x => getter(x).ToString());
		}

		public void RegisterColumnIf(bool expression, string header, Func<TModel, decimal?> getter) {
			RegisterColumnIf(expression, header, x => {
				decimal? value = getter(x);

				return value.HasValue ? string.Format(_culture, "{0:C}", value.Value) : string.Empty;
			});
		}

		public void RegisterColumnIf(bool expression, string header, Func<TModel, DateTime?> getter) {
			RegisterColumnIf(expression, header, x => {
				DateTime? value = getter(x);

				return value.HasValue ? value.Value.ToString(_culture.DateTimeFormat.LongDatePattern, _culture) : string.Empty;
			});
		}

		public void RegisterColumnIf(bool expression, string header, Func<TModel, bool?> getter) {
			RegisterColumnIf(expression, header, x => {
				bool? value = getter(x);
				if (!value.HasValue)
					return string.Empty;

				return value.Value ? "Да" : "Нет";
			});
		}

		public TableConfigurationBuilder<TChildModel> RegisterChild<TChildModel>(string title, Func<TModel, IEnumerable<TChildModel>> getter) {
			TableConfigurationBuilder<TChildModel> conf = new TableConfigurationBuilder<TChildModel>(title, x => {
				if (getter((TModel)x) != null)
					return getter((TModel)x).Cast<object>();

				return new List<object>();
			}, _culture);
			Value.ChildrenMap.Add((ChildTableConfiguration)conf.Value);

			return conf;
		}
	}

	public class TableConfiguration {
		public string Title { get; set; }
		public Dictionary<string, Func<object, string>> ColumnsMap = new Dictionary<string, Func<object, string>>();
		public List<ChildTableConfiguration> ChildrenMap = new List<ChildTableConfiguration>();
	}

	public class ChildTableConfiguration : TableConfiguration {
		public Func<object, IEnumerable<object>> Getter { get; set; }
	}

	public static class TableWriter {
		public static MemoryStream Write<TModel>(ITableWriter writer, IEnumerable<TModel> models, TableConfiguration configuration, bool appendNumsToChildColumns = true) {

			List<string> parentHeader = configuration.ColumnsMap.Keys.ToList();
			Func<object, string>[] parentRow = configuration.ColumnsMap.Values.ToArray();

			ChildTableConfiguration childTable = configuration.ChildrenMap.FirstOrDefault();

			if (childTable == null) {
				writer.WriteHeader(parentHeader.ToArray());

				foreach (TModel model in models) {
					var list = new List<string>();

					foreach (Func<object, string> keyValuePair in parentRow) {
						string row = keyValuePair(model);
						list.Add(row);
					}
					writer.WriteRow(list.ToArray());
					writer.WriteRow();
				}
			}
			else {
				Dictionary<string, Func<object, string>> columnsMap = childTable.ColumnsMap;

				var childHeader = new List<string>();

				foreach (KeyValuePair<string, Func<object, string>> keyValuePair in columnsMap) {
					string s = keyValuePair.Key;
					childHeader.Add(s);
				}

				Func<object, IEnumerable<object>> func = childTable.Getter;

				int max = 0;
				foreach (TModel model in models) {
					IEnumerable<object> cc = func(model);
					max = max > cc.Count() ? max : cc.Count();
				}

				for (int i = 1; i <= max; i++) {
					foreach (string s in childHeader) {
						if (appendNumsToChildColumns) {
							parentHeader.Add(s + i);
						}
						else {
							parentHeader.Add(s);
						}
					}
				}

				writer.WriteHeader(parentHeader.ToArray());

				foreach (TModel model in models) {
					var list = new List<string>();

					foreach (Func<object, string> keyValuePair in parentRow) {
						string row = keyValuePair(model);
						list.Add(row);
					}


					IEnumerable<object> cc = func(model);
					foreach (object c in cc) {
						foreach (KeyValuePair<string, Func<object, string>> keyValuePair in columnsMap) {
							string val = keyValuePair.Value(c);
							list.Add(val);
						}
					}
					writer.WriteRow(list.ToArray());
					writer.WriteRow();
				}
			}

			writer.AutosizeColumns();
			return writer.GetStream();
		}
	}
}
