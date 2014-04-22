using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using FormattedExcelExport.Infrastructure;
using FormattedExcelExport.Style;


namespace FormattedExcelExport.Configuaration {
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
}