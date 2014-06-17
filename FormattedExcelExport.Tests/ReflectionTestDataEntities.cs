using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using FormattedExcelExport.Reflection;

namespace FormattedExcelExport.Tests {
    [ExcelExportClassName(Name = "Тестовые данные")]
    class ReflectionTestDataEntities {
        private readonly List<EnumProp1> _enumField1;
        private readonly List<EnumProp2> _enumField2;
        private readonly List<EnumProp3> _enumField3;
        private readonly List<EnumProp4> _enumField4;
        private readonly List<EnumProp5> _enumField5;

        public ReflectionTestDataEntities() {
            List<Type> generalTypes = GeneralTypes;
            List<PropertyInfo> nonEnumerableProperties = GetType().GetProperties().Where(x => generalTypes.Contains(x.PropertyType)).ToList();
            Random rand = new Random();
	        foreach (PropertyInfo nonEnumerableProperty in nonEnumerableProperties) {
	            Object value;
	            switch (nonEnumerableProperty.PropertyType.Name) {
                    case "String":
                        value = Guid.NewGuid().ToString();
                        nonEnumerableProperty.SetValue(this, Convert.ChangeType(value, nonEnumerableProperty.PropertyType));
                        break;
                    case "Int32":
                        value = rand.Next();
                        nonEnumerableProperty.SetValue(this, Convert.ChangeType(value, nonEnumerableProperty.PropertyType));
                        break;
                    case "Boolean":
                        value = Convert.ToBoolean(rand.Next(2));
                        nonEnumerableProperty.SetValue(this, Convert.ChangeType(value, nonEnumerableProperty.PropertyType));
                        break;
                }
            }
	        _enumField1 = new List<EnumProp1>();
            for (int i = 0; i < rand.Next(4); i++) {
                _enumField1.Add(new EnumProp1());
            }
            _enumField2 = new List<EnumProp2>();
            for (int i = 0; i < rand.Next(4); i++) {
                _enumField2.Add(new EnumProp2());
            }
            _enumField3 = new List<EnumProp3>();
            for (int i = 0; i < rand.Next(4); i++) {
                _enumField3.Add(new EnumProp3());
            }
            _enumField4 = new List<EnumProp4>();
            for (int i = 0; i < rand.Next(4); i++) {
                _enumField4.Add(new EnumProp4());
            }
            _enumField5 = new List<EnumProp5>();
            for (int i = 0; i < rand.Next(4); i++) {
                _enumField5.Add(new EnumProp5());
            }
        }

        [ExcelExport(PropertyName = "Enum Field 1")]
        public List<EnumProp1> EnumField1 {
            get { return _enumField1; }
        }
        [ExcelExport(PropertyName = "Enum Field 2")]
        public List<EnumProp2> EnumField2 {
            get { return _enumField2; }
        }
        [ExcelExport(PropertyName = "Enum Field 3")]
        public List<EnumProp3> EnumField3 {
            get { return _enumField3; }
        }
        [ExcelExport(PropertyName = "Enum Field 4")]
        public List<EnumProp4> EnumField4 {
            get { return _enumField4; }
        }
        [ExcelExport(PropertyName = "Enum Field 5")]
        public List<EnumProp5> EnumField5 {
            get { return _enumField5; }
        }

        [ExcelExport(PropertyName = "Свойство 1")]
        public int Prop1 { get; private set; }

        [ExcelExport(PropertyName = "Свойство 2")]
        public string Prop2 { get; private set; }

        [ExcelExport(PropertyName = "Свойство 3")]
        public bool Prop3 { get; private set; }

        [ExcelExport(PropertyName = "Свойство 4")]
        public int Prop4 { get; private set; }

        [ExcelExport(PropertyName = "Свойство 5")]
        public string Prop5 { get; private set; }

        [ExcelExport(PropertyName = "Свойство 6")]
        public bool Prop6 { get; private set; }

        [ExcelExport(PropertyName = "Свойство 7")]
        public int Prop7 { get; private set; }

        [ExcelExport(PropertyName = "Свойство 8")]
        public string Prop8 { get; private set; }

        [ExcelExport(PropertyName = "Свойство 9")]
        public bool Prop9 { get; private set; }

        [ExcelExport(PropertyName = "Свойство 10")]
        public int Prop10 { get; private set; }

        [ExcelExport(PropertyName = "Свойство 11")]
        public string Prop11 { get; private set; }

        [ExcelExport(PropertyName = "Свойство 12")]
        public bool Prop12 { get; private set; }

        [ExcelExport(PropertyName = "Свойство 13")]
        public int Prop13 { get; private set; }

        [ExcelExport(PropertyName = "Свойство 14")]
        public string Prop14 { get; private set; }

        [ExcelExport(PropertyName = "Свойство 15")]
        public bool Prop15 { get; private set; }

        [ExcelExport(PropertyName = "Свойство 16")]
        public int Prop16 { get; private set; }

        [ExcelExport(PropertyName = "Свойство 17")]
        public string Prop17 { get; private set; }

        [ExcelExport(PropertyName = "Свойство 18")]
        public bool Prop18 { get; private set; }

        [ExcelExport(PropertyName = "Свойство 19")]
        public int Prop19 { get; private set; }

        [ExcelExport(PropertyName = "Свойство 20")]
        public string Prop20 { get; private set; }

        [ExcelExport(PropertyName = "Свойство 21")]
        public bool Prop21 { get; private set; }

        [ExcelExport(PropertyName = "Свойство 22")]
        public int Prop22 { get; private set; }

        [ExcelExport(PropertyName = "Свойство 23")]
        public string Prop23 { get; private set; }

        [ExcelExport(PropertyName = "Свойство 24")]
        public bool Prop24 { get; private set; }

        [ExcelExport(PropertyName = "Свойство 25")]
        public int Prop25 { get; private set; }

        [ExcelExport(PropertyName = "Свойство 26")]
        public string Prop26 { get; private set; }

        [ExcelExport(PropertyName = "Свойство 27")]
        public bool Prop27 { get; private set; }

        [ExcelExport(PropertyName = "Свойство 28")]
        public int Prop28 { get; private set; }

        [ExcelExport(PropertyName = "Свойство 29")]
        public string Prop29 { get; private set; }

        [ExcelExport(PropertyName = "Свойство 30")]
        public bool Prop30 { get; private set; }

        [ExcelExportClassName(Name = "Перечисл. свойство1")]
        public sealed class EnumProp1 {
            public EnumProp1() {
                List<Type> generalTypes = GeneralTypes;
                List<PropertyInfo> nonEnumerableProperties = GetType().GetProperties().Where(x => generalTypes.Contains(x.PropertyType)).ToList();
                Random rand = new Random();
	            foreach (PropertyInfo nonEnumerableProperty in nonEnumerableProperties) {
	                Object value;
	                switch (nonEnumerableProperty.PropertyType.Name) {
                        case "String":
                            value = Guid.NewGuid().ToString();
                            nonEnumerableProperty.SetValue(this, Convert.ChangeType(value, nonEnumerableProperty.PropertyType));
                            break;
                        case "Int32":
                            value = rand.Next();
                            nonEnumerableProperty.SetValue(this, Convert.ChangeType(value, nonEnumerableProperty.PropertyType));
                            break;
                        case "Boolean":
                            value = Convert.ToBoolean(rand.Next(2));
                            nonEnumerableProperty.SetValue(this, Convert.ChangeType(value, nonEnumerableProperty.PropertyType));
                            break;
                    }
                }
            }

            [ExcelExport(PropertyName = "Поле1")]
            public string Field1 { get; private set; }

            [ExcelExport(PropertyName = "Поле2")]
            public int Field2 { get; private set; }

            [ExcelExport(PropertyName = "Поле3")]
            public bool Field3 { get; private set; }

            [ExcelExport(PropertyName = "Поле4")]
            public string Field4 { get; private set; }

            [ExcelExport(PropertyName = "Поле5")]
            public int Field5 { get; private set; }

            [ExcelExport(PropertyName = "Поле6")]
            public bool Field6 { get; private set; }
        }

        [ExcelExportClassName(Name = "Перечисл. свойство2")]
        public sealed class EnumProp2 {
            public EnumProp2() {
                List<Type> generalTypes = GeneralTypes;
                List<PropertyInfo> nonEnumerableProperties = GetType().GetProperties().Where(x => generalTypes.Contains(x.PropertyType)).ToList();
                Random rand = new Random();
	            foreach (PropertyInfo nonEnumerableProperty in nonEnumerableProperties) {
	                Object value;
	                switch (nonEnumerableProperty.PropertyType.Name) {
                        case "String":
                            value = Guid.NewGuid().ToString();
                            nonEnumerableProperty.SetValue(this, Convert.ChangeType(value, nonEnumerableProperty.PropertyType));
                            break;
                        case "Int32":
                            value = rand.Next();
                            nonEnumerableProperty.SetValue(this, Convert.ChangeType(value, nonEnumerableProperty.PropertyType));
                            break;
                        case "Boolean":
                            value = Convert.ToBoolean(rand.Next(2));
                            nonEnumerableProperty.SetValue(this, Convert.ChangeType(value, nonEnumerableProperty.PropertyType));
                            break;
                    }
                }
            }

            [ExcelExport(PropertyName = "Поле1")]
            public string Field1 { get; private set; }

            [ExcelExport(PropertyName = "Поле2")]
            public int Field2 { get; private set; }

            [ExcelExport(PropertyName = "Поле3")]
            public bool Field3 { get; private set; }

            [ExcelExport(PropertyName = "Поле4")]
            public string Field4 { get; private set; }

            [ExcelExport(PropertyName = "Поле5")]
            public int Field5 { get; private set; }

            [ExcelExport(PropertyName = "Поле6")]
            public bool Field6 { get; private set; }

            [ExcelExport(PropertyName = "Поле7")]
            public string Field7 { get; private set; }

            [ExcelExport(PropertyName = "Поле8")]
            public int Field8 { get; private set; }

            [ExcelExport(PropertyName = "Поле9")]
            public bool Field9 { get; private set; }

            [ExcelExport(PropertyName = "Поле10")]
            public string Field10 { get; private set; }
        }

        [ExcelExportClassName(Name = "Перечисл. свойство3")]
        public sealed class EnumProp3 {
            public EnumProp3() {
                List<Type> generalTypes = GeneralTypes;
                List<PropertyInfo> nonEnumerableProperties = GetType().GetProperties().Where(x => generalTypes.Contains(x.PropertyType)).ToList();
                Random rand = new Random();
	            foreach (PropertyInfo nonEnumerableProperty in nonEnumerableProperties) {
	                Object value;
	                switch (nonEnumerableProperty.PropertyType.Name) {
                        case "String":
                            value = Guid.NewGuid().ToString();
                            nonEnumerableProperty.SetValue(this, Convert.ChangeType(value, nonEnumerableProperty.PropertyType));
                            break;
                        case "Int32":
                            value = rand.Next();
                            nonEnumerableProperty.SetValue(this, Convert.ChangeType(value, nonEnumerableProperty.PropertyType));
                            break;
                        case "Boolean":
                            value = Convert.ToBoolean(rand.Next(2));
                            nonEnumerableProperty.SetValue(this, Convert.ChangeType(value, nonEnumerableProperty.PropertyType));
                            break;
                    }
                }
            }

            [ExcelExport(PropertyName = "Поле1")]
            public string Field1 { get; private set; }

            [ExcelExport(PropertyName = "Поле2")]
            public int Field2 { get; private set; }

            [ExcelExport(PropertyName = "Поле3")]
            public bool Field3 { get; private set; }

            [ExcelExport(PropertyName = "Поле4")]
            public string Field4 { get; private set; }

            [ExcelExport(PropertyName = "Поле5")]
            public int Field5 { get; private set; }

            [ExcelExport(PropertyName = "Поле6")]
            public bool Field6 { get; private set; }

            [ExcelExport(PropertyName = "Поле7")]
            public string Field7 { get; private set; }

            [ExcelExport(PropertyName = "Поле8")]
            public int Field8 { get; private set; }

            [ExcelExport(PropertyName = "Поле9")]
            public bool Field9 { get; private set; }

            [ExcelExport(PropertyName = "Поле10")]
            public string Field10 { get; private set; }
        }

        [ExcelExportClassName(Name = "Перечисл. свойство4")]
        public sealed class EnumProp4 {
            public EnumProp4() {
                List<Type> generalTypes = GeneralTypes;
                List<PropertyInfo> nonEnumerableProperties = GetType().GetProperties().Where(x => generalTypes.Contains(x.PropertyType)).ToList();
                Random rand = new Random();
	            foreach (PropertyInfo nonEnumerableProperty in nonEnumerableProperties) {
	                Object value;
	                switch (nonEnumerableProperty.PropertyType.Name) {
                        case "String":
                            value = Guid.NewGuid().ToString();
                            nonEnumerableProperty.SetValue(this, Convert.ChangeType(value, nonEnumerableProperty.PropertyType));
                            break;
                        case "Int32":
                            value = rand.Next();
                            nonEnumerableProperty.SetValue(this, Convert.ChangeType(value, nonEnumerableProperty.PropertyType));
                            break;
                        case "Boolean":
                            value = Convert.ToBoolean(rand.Next(2));
                            nonEnumerableProperty.SetValue(this, Convert.ChangeType(value, nonEnumerableProperty.PropertyType));
                            break;
                    }
                }
            }

            [ExcelExport(PropertyName = "Поле1")]
            public string Field1 { get; private set; }

            [ExcelExport(PropertyName = "Поле2")]
            public int Field2 { get; private set; }

            [ExcelExport(PropertyName = "Поле3")]
            public bool Field3 { get; private set; }

            [ExcelExport(PropertyName = "Поле4")]
            public string Field4 { get; private set; }

            [ExcelExport(PropertyName = "Поле5")]
            public int Field5 { get; private set; }

            [ExcelExport(PropertyName = "Поле6")]
            public bool Field6 { get; private set; }

            [ExcelExport(PropertyName = "Поле7")]
            public string Field7 { get; private set; }

            [ExcelExport(PropertyName = "Поле8")]
            public int Field8 { get; private set; }
        }

        [ExcelExportClassName(Name = "Перечисл. свойство5")]
        public sealed class EnumProp5 {
            public EnumProp5() {
                List<Type> generalTypes = GeneralTypes;
                List<PropertyInfo> nonEnumerableProperties = GetType().GetProperties().Where(x => generalTypes.Contains(x.PropertyType)).ToList();
                Random rand = new Random();
	            foreach (PropertyInfo nonEnumerableProperty in nonEnumerableProperties) {
	                Object value;
	                switch (nonEnumerableProperty.PropertyType.Name) {
                        case "String":
                            value = Guid.NewGuid().ToString();
                            nonEnumerableProperty.SetValue(this, Convert.ChangeType(value, nonEnumerableProperty.PropertyType));
                            break;
                        case "Int32":
                            value = rand.Next();
                            nonEnumerableProperty.SetValue(this, Convert.ChangeType(value, nonEnumerableProperty.PropertyType));
                            break;
                        case "Boolean":
                            value = Convert.ToBoolean(rand.Next(2));
                            nonEnumerableProperty.SetValue(this, Convert.ChangeType(value, nonEnumerableProperty.PropertyType));
                            break;
                    }
                }
            }

            [ExcelExport(PropertyName = "Поле1")]
            public string Field1 { get; private set; }

            [ExcelExport(PropertyName = "Поле2")]
            public int Field2 { get; private set; }

            [ExcelExport(PropertyName = "Поле3")]
            public bool Field3 { get; private set; }

            [ExcelExport(PropertyName = "Поле4")]
            public string Field4 { get; private set; }

            [ExcelExport(PropertyName = "Поле5")]
            public int Field5 { get; private set; }

            [ExcelExport(PropertyName = "Поле6")]
            public bool Field6 { get; private set; }

            [ExcelExport(PropertyName = "Поле7")]
            public string Field7 { get; private set; }

            [ExcelExport(PropertyName = "Поле8")]
            public int Field8 { get; private set; }

            [ExcelExport(PropertyName = "Поле9")]
            public bool Field9 { get; private set; }

            [ExcelExport(PropertyName = "Поле10")]
            public string Field10 { get; private set; }
        }

        private static List<Type> GeneralTypes {
            get {
                List<Type> generalTypes = new List<Type> {
                    typeof (string),
                    typeof (DateTime),
                    typeof (decimal),
                    typeof (int),
                    typeof (bool)
                };
                return generalTypes;
            }
        }
    }
}
