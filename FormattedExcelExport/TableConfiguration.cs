using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace FormattedExcelExport {
	public class TableConfiguration {
		public string Title { get; set; }
		public readonly Dictionary<string, Func<object, string>> ColumnsMap = new Dictionary<string, Func<object, string>>();
		public readonly List<ChildTableConfiguration> ChildrenMap = new List<ChildTableConfiguration>();
	}

	public class ChildTableConfiguration : TableConfiguration {
		public Func<object, IEnumerable<object>> Getter { get; set; }
	}



	public sealed class TableConfigurationBuilder<TModel> {
		public class ConditionTheme {
			private TableWriterStyle _style;
			private Func<TModel, bool> _condition;
			public ConditionTheme(TableWriterStyle style, Func<TModel, bool> condition) {
				Style = style;
				Condition = condition;
			}
			public Func<TModel, bool> Condition {
				get { return _condition; }
				set { _condition = value; }
			}
			public TableWriterStyle Style {
				get { return _style; }
				set { _style = value; }
			}
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

		public void RegisterColumn(string header, Func<TModel, string> getter) {
			Value.ColumnsMap.Add(header, x => getter((TModel)x));
		}

		public void RegisterColumn(string header, Func<TModel, decimal?> getter) {
			RegisterColumn(header, x => {
				decimal? value = getter(x);
				return value.HasValue ? string.Format(_culture, "{0:C}", value.Value) : string.Empty;
			});
		}

		public void RegisterColumnIf(bool expression, string header, Func<TModel, string> getter, ConditionTheme conditionTheme = null) {
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
			TableConfigurationBuilder<TChildModel> tableConfigurationBuilder = new TableConfigurationBuilder<TChildModel>(title, x => {
				if (getter((TModel)x) != null)
					return getter((TModel)x).Cast<object>();

				return new List<object>();
			}, _culture);
			Value.ChildrenMap.Add((ChildTableConfiguration)tableConfigurationBuilder.Value);

			return tableConfigurationBuilder;
		}
	}
}
