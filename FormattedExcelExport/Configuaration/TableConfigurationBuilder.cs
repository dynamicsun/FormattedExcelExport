using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using FormattedExcelExport.Infrastructure;
using FormattedExcelExport.Style;
using NPOI.HSSF.Record.Chart;


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

	    private readonly Func<dynamic, dynamic> _getter = model => {
	        dynamic value = null;
	        if (model is decimal? || model is int? || model is float?) value = model != null ? (double?) Convert.ToDouble(model) : null;
	        if (model is DateTime? || model is string) value = model;
	        if (model is bool?) value = model != null ? (model == true ? "Да" : "Нет") : string.Empty;
	        return value;
	    };
		public TableConfigurationBuilder(string title, CultureInfo culture) {
			Value = new TableConfiguration { Title = title };
			_culture = culture;
		}
        public TableConfigurationBuilder(string title, Func<object, IEnumerable<object>> getter, CultureInfo culture) {
            Value = new ChildTableConfiguration { Getter = getter, Title = title };
            _culture = culture;
        }
        public void RegisterColumn(string header, Func<TModel, string> getter, ConditionTheme conditionTheme = null) {
            RegisterColumnIf(header, getter, conditionTheme);
        }
        public void RegisterColumn(string header, Func<TModel, int?> getter, ConditionTheme conditionTheme = null) {
            RegisterColumnIf(header, x => _getter(getter(x)), conditionTheme);
        }
        public void RegisterColumn(string header, Func<TModel, decimal?> getter, ConditionTheme conditionTheme = null) {
            RegisterColumnIf(header, x => _getter(getter(x)), conditionTheme);
        }
        public void RegisterColumn(string header, Func<TModel, float?> getter, ConditionTheme conditionTheme = null) {
            RegisterColumnIf(header, x => _getter(getter(x)), conditionTheme);
        }
        public void RegisterColumn(string header, Func<TModel, DateTime?> getter, ConditionTheme conditionTheme = null) {
            RegisterColumnIf(header, x => _getter(getter(x)), conditionTheme);
        }
        public void RegisterColumn(string header, Func<TModel, bool?> getter, ConditionTheme conditionTheme = null) {
            RegisterColumnIf(header, x => _getter(getter(x)), conditionTheme);
        }
		public void RegisterColumnIf(string header, Func<TModel, dynamic> getter, ConditionTheme conditionTheme = null) {
			if(conditionTheme == null)
				conditionTheme = new ConditionTheme(null, x => false);
			Value.ColumnsMap.Add(header, new AggregatedContainer(x => getter((TModel)x), y => conditionTheme.Condition((TModel)y), conditionTheme.Style));
		}
		public TableConfigurationBuilder<TChildModel> RegisterChild<TChildModel>(string title, Func<TModel, IEnumerable<TChildModel>> getter) {
			var tableConfigurationBuilder = new TableConfigurationBuilder<TChildModel>(title, x => {
				if (getter((TModel)x) != null)
					return getter((TModel)x).Cast<object>();

				return new List<object>();
			}, _culture);
			Value.ChildrenMap.Add((ChildTableConfiguration)tableConfigurationBuilder.Value);

			return tableConfigurationBuilder;
		}
	}
}