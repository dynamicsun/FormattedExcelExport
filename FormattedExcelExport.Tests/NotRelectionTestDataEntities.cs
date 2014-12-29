using System;
using System.Collections.Generic;
using System.Globalization;
using FormattedExcelExport.Configuaration;
using FormattedExcelExport.Reflection;
using FormattedExcelExport.Style;

namespace FormattedExcelExport.Tests {
    internal static class NotRelectionTestDataEntities {
        internal static TestData CreateSimpleTestData(bool style = false) {
            var dataStructure = CreateSimpleTestDataConfigurationBuilder(style);
            var data = CreateSimpleTestDataModels();
            return new TestData(dataStructure, data);
        }

        internal static TestData CreateSimpleTestRowOverflowData(bool style = false) {
            var dataStructure = CreateSimpleTestDataConfigurationBuilder(style);
            var data = CreateTestRowOverflowDataModels();
            return new TestData(dataStructure, data);
        }
        private static TableConfigurationBuilder<ClientExampleModel> CreateSimpleTestDataConfigurationBuilder(bool style = false) {
            if (style) {
                var condStyle = new TableWriterStyle();
                condStyle.RegularCell.BackgroundColor = new AdHocCellStyle.Color(255, 0, 0);
                var condStyle2 = new TableWriterStyle();
                condStyle2.RegularCell.BackgroundColor = new AdHocCellStyle.Color(0, 255, 0);
                var condStyle3 = new TableWriterStyle();
                condStyle3.RegularChildCell.BackgroundColor = new AdHocCellStyle.Color(0, 0, 255);
                

                var confBuilder = new TableConfigurationBuilder<ClientExampleModel>("Клиент", new CultureInfo("ru-RU"));
                confBuilder.RegisterColumn("Название", x => x.Title, new TableConfigurationBuilder<ClientExampleModel>.ConditionTheme(condStyle, x => x.Title == "Вторая компания"));
                confBuilder.RegisterColumn("Дата регистрации", x => x.RegistrationDate);
                confBuilder.RegisterColumn("Телефон", x => x.Phone, new TableConfigurationBuilder<ClientExampleModel>.ConditionTheme(condStyle2, x => x.Okato == "OPEEHBSSDD"));
                confBuilder.RegisterColumn("ИНН", x => x.Inn);
                confBuilder.RegisterColumn("Окато", x => x.Okato);
                confBuilder.RegisterColumn("Доход", x => x.Revenue);
                confBuilder.RegisterColumn("Число сотрудников", x => x.EmployeeCount); 
                confBuilder.RegisterColumn("Активен", x => x.IsActive);
                confBuilder.RegisterColumn("Свойство 1", x => x.Prop1);
                confBuilder.RegisterColumn("Свойство 2", x => x.Prop2);
                confBuilder.RegisterColumn("Свойство 3", x => x.Prop3);
                confBuilder.RegisterColumn("Свойство 4", x => x.Prop4);
                confBuilder.RegisterColumn("Свойство 5", x => x.Prop5);
                confBuilder.RegisterColumn("Свойство 6", x => x.Prop6);
                confBuilder.RegisterColumn("Свойство 7", x => x.Prop7);
                var contact = confBuilder.RegisterChild("Контакт", x => x.Contacts);
                contact.RegisterColumn("Название", x => x.Title, new TableConfigurationBuilder<ClientExampleModel.Contact>.ConditionTheme(condStyle3, x => x.Title.StartsWith("О")));
                contact.RegisterColumn("Email", x => x.Email);

                var contract = confBuilder.RegisterChild("Контракт", x => x.Contracts);
                contract.RegisterColumn("Дата начала", x => x.BeginDate);
                contract.RegisterColumn("Дата окончания", x => x.EndDate);
                contract.RegisterColumn("Статус222222222222222222222", x => x.Status, new TableConfigurationBuilder<ClientExampleModel.Contract>.ConditionTheme(new TableWriterStyle(), x => true));

                var product = confBuilder.RegisterChild("Продукт", x => x.Products);
                product.RegisterColumn("Наименование", x => x.Title);
                product.RegisterColumn("Количество", x => x.Amount);

                var enumProp1 = confBuilder.RegisterChild("Переч. свойство 1", x => x.EnumProps1);
                enumProp1.RegisterColumn("Поле2222222222222222222222 1-", x => x.Field1);
                enumProp1.RegisterColumn("Поле 2-", x => x.Field2);
                enumProp1.RegisterColumn("Поле22222222222222 3-", x => x.Field3);

                var enumProp2 = confBuilder.RegisterChild("Перечисл. свойство 2", x => x.EnumProps2);
                enumProp2.RegisterColumn("Поле 4-", x => x.Field4);
                enumProp2.RegisterColumn("Поле 5-", x => x.Field5);
                enumProp2.RegisterColumn("Поле 6-", x => x.Field6);
                enumProp2.RegisterColumn("Поле 72222222222222222222222222222222222-", x => x.Field7);
                enumProp2.RegisterColumn("Поле 822222222222222222222222222222-", x => x.Field8);
                return confBuilder;
            }
            else {
                var confBuilder = new TableConfigurationBuilder<ClientExampleModel>("Клиент", new CultureInfo("ru-RU"));
                confBuilder.RegisterColumn("Названиеd", x => x.Title);
                confBuilder.RegisterColumn("Дата регистрации", x => x.RegistrationDate);
                confBuilder.RegisterColumn("Телефон", x => x.Phone);
                confBuilder.RegisterColumn("ИНН", x => x.Inn);
                confBuilder.RegisterColumn("Окато", x => x.Okato);
                confBuilder.RegisterColumn("Доход", x => x.Revenue);
                confBuilder.RegisterColumn("Число сотрудников", x => x.EmployeeCount);
                confBuilder.RegisterColumn("Активен", x => x.IsActive);
                confBuilder.RegisterColumn("Свойство 1", x => x.Prop1);
                confBuilder.RegisterColumn("Свойство 2", x => x.Prop2);
                confBuilder.RegisterColumn("Свойство 3", x => x.Prop3);
                confBuilder.RegisterColumn("Свойство 3333333333333333333333333333334", x => x.Prop4);
                confBuilder.RegisterColumn("Свойство 5", x => x.Prop5);
                confBuilder.RegisterColumn("Свойство 6", x => x.Prop6);
                confBuilder.RegisterColumn("Свойство 7", x => x.Prop7);

                var contact = confBuilder.RegisterChild("Контакт", x => x.Contacts);
                contact.RegisterColumn("Название3333333333333333333333333333", x => x.Title);
                contact.RegisterColumn("Email", x => x.Email);

                var contract = confBuilder.RegisterChild("Контракт", x => x.Contracts);
                contract.RegisterColumn("Дата начала", x => x.BeginDate);
                contract.RegisterColumn("Дата окончания", x => x.EndDate);
                contract.RegisterColumn("Статус", x => x.Status);

                var product = confBuilder.RegisterChild("Продукт", x => x.Products);
                product.RegisterColumn("Наименование", x => x.Title);
                product.RegisterColumn("Количество", x => x.Amount);

                var enumProp1 = confBuilder.RegisterChild("Переч. свойство 1", x => x.EnumProps1);
                enumProp1.RegisterColumn("Поле 1", x => x.Field1);
                enumProp1.RegisterColumn("Поле 2", x => x.Field2);
                enumProp1.RegisterColumn("Поле 3", x => x.Field3);

                var enumProp2 = confBuilder.RegisterChild("Перечисл. свойство 2", x => x.EnumProps2);
                enumProp2.RegisterColumn("Поле 4", x => x.Field4);
                enumProp2.RegisterColumn("Поле 5", x => x.Field5);
                enumProp2.RegisterColumn("Поле 6", x => x.Field6);
                enumProp2.RegisterColumn("Поле 7", x => x.Field7);
                enumProp2.RegisterColumn("Поле 811111111111111eeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeee", x => x.Field8);
                return confBuilder;
            }
        }

        internal static List<ClientExampleModel> CreateTestRowOverflowDataModels() {
            var models = new List<ClientExampleModel>();
            for (var i = 0; i < 17500; i++) {
                models.Add(new ClientExampleModel(
                    null,
                    DateTime.Now,
                    "+7 333 4442 00",
                    "9040043234562",
                    "OPEEHBSSDD",
                    2352666,
                    336,
                    true,
                    1234,
                    "kdkdkd",
                    true,
                    4321,
                    "ddd",
                    false,
                    8979,
                    new List<ClientExampleModel.Contact> {
						new ClientExampleModel.Contact("Ольга", "olga@mail.ru"),
						new ClientExampleModel.Contact("Иван", "ivan@mail.ru")
					},
                    new List<ClientExampleModel.Contract> {
						new ClientExampleModel.Contract(new DateTime(1999, 1, 7), new DateTime(2009, 11, 9), false),
						new ClientExampleModel.Contract(new DateTime(1989, 2, 4), DateTime.Now, true)
					},
                    new List<ClientExampleModel.Product> {
						new ClientExampleModel.Product("Картофель", 20),
						new ClientExampleModel.Product("Лук", 100)
					},
                    new List<ClientExampleModel.EnumProp1> {
                        new ClientExampleModel.EnumProp1("dsdsd", 432, true),
                        new ClientExampleModel.EnumProp1("dd", 36, true),
                        new ClientExampleModel.EnumProp1("yap", 310, false),
                    },
                    new List<ClientExampleModel.EnumProp2> {
                        new ClientExampleModel.EnumProp2("s", 1, true, "b", 3),
                        new ClientExampleModel.EnumProp2("t", 2, false, "a", 5),
                        new ClientExampleModel.EnumProp2("d", 3, true, "d", 6),
                        new ClientExampleModel.EnumProp2("s", 6, true, "ds", 7),
                        new ClientExampleModel.EnumProp2("yaoo", 7, false, "d", 5)
                    }));
                models.Add(new ClientExampleModel(
                    "Вторая компания",
                    DateTime.Now,
                    "+7 222 1124 44",
                    "5953043385461",
                    "JsKSLPKKHSS",
                    599988,
                    59,
                    false,
                    1234,
                    "kdkdkd",
                    true,
                    4321,
                    "ddd",
                    false,
                    8979,
                    new List<ClientExampleModel.Contact> {
                        new ClientExampleModel.Contact("Олег", "oleg@mail.ru"),
                        new ClientExampleModel.Contact("Анна", ""),
                        new ClientExampleModel.Contact("Николай", "nikolay@mail.ru")
                    },
                    new List<ClientExampleModel.Contract> {
                        new ClientExampleModel.Contract(new DateTime(1999, 1, 7), new DateTime(2007, 3, 22), true),
                        new ClientExampleModel.Contract(new DateTime(1989, 2, 4), new DateTime(2012, 11, 20), false)
                    },
                    new List<ClientExampleModel.Product>(),
                    new List<ClientExampleModel.EnumProp1> {
                        new ClientExampleModel.EnumProp1("dsdsd", 432, true),
                        new ClientExampleModel.EnumProp1("dd", 34, false),
                        new ClientExampleModel.EnumProp1("yap", 3010, false),
                    },
                    new List<ClientExampleModel.EnumProp2> {
                        new ClientExampleModel.EnumProp2("ss", 1, true, "bb", 43),
                        new ClientExampleModel.EnumProp2("tr", 2, false, "aa", 45),
                        new ClientExampleModel.EnumProp2("ds", 3, true, "dd", 46),
                        new ClientExampleModel.EnumProp2("ds", 4, true, "dsds", 47),
                        new ClientExampleModel.EnumProp2("yahoo", 5, false, "d", 45)
                    }));
                models.Add(new ClientExampleModel(
                    "Третья компания",
                    DateTime.Now,
                    "+7 222 1124 44",
                    "5953043385461",
                    "JsKSLPKKHSS",
                    599988,
                    59,
                    false,
                    1234,
                    "kdkdkd",
                    true,
                    4321,
                    "ddd",
                    false,
                    8979,
                    new List<ClientExampleModel.Contact> {
						new ClientExampleModel.Contact("Олег", "oleg@mail.ru"),
                        new ClientExampleModel.Contact("Даниил", "daniil@mail.ru"),
						new ClientExampleModel.Contact("Анна", "anna@anna.anna")
					},
                    new List<ClientExampleModel.Contract> {
						new ClientExampleModel.Contract(new DateTime(1999, 1, 7), new DateTime(2007, 3, 22), true)
					},
                    new List<ClientExampleModel.Product>(),
                    new List<ClientExampleModel.EnumProp1> {
                        new ClientExampleModel.EnumProp1("dsdsd", 432, true),
                        new ClientExampleModel.EnumProp1("dd", 34, false),
                        new ClientExampleModel.EnumProp1("yap", 3010, false),
                    },
                    new List<ClientExampleModel.EnumProp2> {
                        new ClientExampleModel.EnumProp2("sdsdsds", 14, false, "fdfbb", 43),
                        new ClientExampleModel.EnumProp2("tsdsdr", 22, false, "aafdf", 45),
                        new ClientExampleModel.EnumProp2("dsdsds", 35, true, "ddfsd", 46),
                        new ClientExampleModel.EnumProp2("ddagdfs", 46, true, "dsdfsds", 47),
                        new ClientExampleModel.EnumProp2("yahovdfo", 65, false, "dfsdf", 45)
                    }));
            }
            return models; 
        }

        internal static List<ClientExampleModel> CreateSimpleTestDataModels() {
            return new List<ClientExampleModel> {
				new ClientExampleModel(
					"dddddddddddd", 
					DateTime.Now, 
					"+7 333 4442 00", 
					"9040043234562",
					"OPEEHBSSDD",
					2352666,
					336,
					true,
                    1234,
                    "kdkdkd",
                    true,
                    4321,
                    "ddd",
                    false,
                    8979,
					new List<ClientExampleModel.Contact> {
						new ClientExampleModel.Contact("Ольга", "olga@mail.ru"),
						new ClientExampleModel.Contact("Иван", "ivan@mail.ru")
					},
					new List<ClientExampleModel.Contract> {
						new ClientExampleModel.Contract(new DateTime(1999, 1, 7), new DateTime(2009, 11, 9), false),
						new ClientExampleModel.Contract(new DateTime(1989, 2, 4), DateTime.Now, true)
					},
					new List<ClientExampleModel.Product> {
						new ClientExampleModel.Product("Картофель", 20),
						new ClientExampleModel.Product("Лук", 100)
					},
                    new List<ClientExampleModel.EnumProp1> {
                        new ClientExampleModel.EnumProp1("dsdsd", 432, true),
                        new ClientExampleModel.EnumProp1("dd", 36, true),
                        new ClientExampleModel.EnumProp1("yap", 310, false),
                    }, 
                    new List<ClientExampleModel.EnumProp2> {
                        new ClientExampleModel.EnumProp2("s", 1, true, "b", 3),
                        new ClientExampleModel.EnumProp2("t", 2, false, "a", 5),
                        new ClientExampleModel.EnumProp2("d", 3, true, "d", 6),
                        new ClientExampleModel.EnumProp2("s", 6, true, "ds", 7),
                        new ClientExampleModel.EnumProp2("yaoo", 7, false, "d", 5)
                    }),
				new ClientExampleModel(
					"Вторая компания", 
					DateTime.Now, 
					"+7 222 1124 44", 
					"5953043385461",
					"JsKSLPKKHSS",
					599988,
					59,
					false,
                    1234,
                    "kdkdkd",
                    true,
                    4321,
                    "ddd",
                    false,
                    8979,
					new List<ClientExampleModel.Contact> {
						new ClientExampleModel.Contact("Олег", "oleg@mail.ru"),
						new ClientExampleModel.Contact("Анна", ""),
						new ClientExampleModel.Contact("Николай", "nikolay@mail.ru")
					},
					new List<ClientExampleModel.Contract> {
						new ClientExampleModel.Contract(new DateTime(1999, 1, 7), new DateTime(2007, 3, 22), true),
						new ClientExampleModel.Contract(new DateTime(1989, 2, 4), new DateTime(2012, 11, 20), false)
					},
					new List<ClientExampleModel.Product> (),
                    new List<ClientExampleModel.EnumProp1> {
                        new ClientExampleModel.EnumProp1("dsdsd", 432, true),
                        new ClientExampleModel.EnumProp1("dd", 34, false),
                        new ClientExampleModel.EnumProp1("yap", 3010, false),
                    }, 
                    new List<ClientExampleModel.EnumProp2> {
                        new ClientExampleModel.EnumProp2("ss", 1, true, "bb", 43),
                        new ClientExampleModel.EnumProp2("tr", 2, false, "aa", 45),
                        new ClientExampleModel.EnumProp2("ds", 3, true, "dd", 46),
                        new ClientExampleModel.EnumProp2("ds", 4, true, "dsds", 47),
                        new ClientExampleModel.EnumProp2("yahoo", 5, false, "d", 45)
                    }),
                new ClientExampleModel(
					"Третья компания", 
					DateTime.Now, 
					"+7 222 1124 44", 
					"5953043385461",
					"JsKSLPKKHSS",
					599988,
					59,
					false,
                    1234,
                    "kdkdkd",
                    true,
                    4321,
                    "ddd",
                    false,
                    8979,
					new List<ClientExampleModel.Contact> {
						new ClientExampleModel.Contact("Олег", "oleg@mail.ru"),
                        new ClientExampleModel.Contact("Даниил", "daniil@mail.ru"),
						new ClientExampleModel.Contact("Анна", "anna@anna.anna")
					},
					new List<ClientExampleModel.Contract> {
						new ClientExampleModel.Contract(new DateTime(1999, 1, 7), new DateTime(2007, 3, 22), true)
					},
					new List<ClientExampleModel.Product> (),
                    new List<ClientExampleModel.EnumProp1> {
                        new ClientExampleModel.EnumProp1("dsdsd", 432, true),
                        new ClientExampleModel.EnumProp1("dd", 34, false),
                        new ClientExampleModel.EnumProp1("yap", 3010, false),
                    }, 
                    new List<ClientExampleModel.EnumProp2> {
                        new ClientExampleModel.EnumProp2("sdsdsds", 14, false, "fdfbb", 43),
                        new ClientExampleModel.EnumProp2("tsdsdr", 22, false, "aafdf", 45),
                        new ClientExampleModel.EnumProp2("dsdsds", 35, true, "ddfsd", 46),
                        new ClientExampleModel.EnumProp2("ddagdfs", 46, true, "dsdfsds", 47),
                        new ClientExampleModel.EnumProp2("yahovdfo", 65, false, "dfsdf", 45)
                    }),
                new ClientExampleModel(
					"Четвертая компания", 
					DateTime.Now, 
					"+7 234 4324 44", 
					"5953043385461",
					"JsKSLPKKHSS",
					599988,
					59,
					false,
                    1234,
                    "kdfsfsdkd",
                    true,
                    4321,
                    "ddasd",
                    false,
                    8979,
					new List<ClientExampleModel.Contact> {
						new ClientExampleModel.Contact("Олег", "oleg@mail.ru"),
						new ClientExampleModel.Contact("Анна", ""),
						new ClientExampleModel.Contact("Николай", "nikolay@mail.ru")
					},
					new List<ClientExampleModel.Contract> (),
					new List<ClientExampleModel.Product> (),
                    new List<ClientExampleModel.EnumProp1> {
                        new ClientExampleModel.EnumProp1("dsdsd", 432, true),
                        new ClientExampleModel.EnumProp1("dd", 34, false),
                        new ClientExampleModel.EnumProp1("yap", 3010, false),
                    }, 
                    new List<ClientExampleModel.EnumProp2> {
                        new ClientExampleModel.EnumProp2("sds", 1, true, "bab", 43),
                        new ClientExampleModel.EnumProp2("tfr", 2, false, "ssaa", 45),
                        new ClientExampleModel.EnumProp2("dds", 3, true, "dssd", 46),
                        new ClientExampleModel.EnumProp2("dfs", 4, true, "dsssds", 47),
                        new ClientExampleModel.EnumProp2("yahoo", 5, false, "ssd", 45)
                    }),
                new ClientExampleModel(
					"Пятая компания", 
					DateTime.Now, 
					"+7 234 5444 44", 
					"6763043355661",
					"JsKSLPKKHSS",
					5988,
					5,
					false,
                    1234,
                    "kdsdkd",
                    true,
                    4321,
                    "asd",
                    false,
                    8979,
					new List<ClientExampleModel.Contact> (),
					new List<ClientExampleModel.Contract> (),
					new List<ClientExampleModel.Product> (),
                    new List<ClientExampleModel.EnumProp1> {
                        new ClientExampleModel.EnumProp1("dsdsd", 43432, true),
                        new ClientExampleModel.EnumProp1("dd", 34, false),
                        new ClientExampleModel.EnumProp1("yap", 301430, false),
                    }, 
                    new List<ClientExampleModel.EnumProp2> {
                        new ClientExampleModel.EnumProp2("sвыds", 1, true, "bdasdab", 433),
                        new ClientExampleModel.EnumProp2("tвыыfr", 2, false, "ssdgaa", 4435),
                        new ClientExampleModel.EnumProp2("dвыывds", 3, true, "dsgfsd", 46),
                        new ClientExampleModel.EnumProp2("dыывfs", 4, true, "dssfdssds", 4347),
                        new ClientExampleModel.EnumProp2("yaвыыhoo", 5, false, "sgdsd", 44345)
                    })
			};
        }
        [ExcelExportClassName(Name = "Клиент")]
        public class ClientExampleModel {
            private readonly string _title;
            private readonly DateTime _registrationDate;
            private readonly string _phone;
            private readonly string _inn;
            private readonly string _okato;
            private readonly decimal _revenue;
            private readonly int _employeeCount;
            private readonly bool _isActive;
            private readonly int _prop1;
            private readonly string _prop2;
            private readonly bool _prop3;
            private readonly int _prop4;
            private readonly string _prop5;
            private readonly bool _prop6;
            private readonly int _prop7;
            private readonly List<Contact> _contacts;
            private readonly List<Contract> _contracts;
            private readonly List<Product> _products;
            private readonly List<EnumProp1> _enumProp1; 
            private readonly List<EnumProp2> _enumProp2;

            public ClientExampleModel(string title, DateTime registrationDate, string phone, string inn, string okato, decimal revenue, int employeeCount, bool isActive, int prop1, string prop2, 
                bool prop3, int prop4, string prop5, bool prop6, int prop7, List<Contact> contacts, List<Contract> contracts, List<Product> products, List<EnumProp1> enumProp1, List<EnumProp2> enumProp2 ) {
                _title = title;
                _registrationDate = registrationDate;
                _phone = phone;
                _inn = inn;
                _okato = okato;
                _revenue = revenue;
                _employeeCount = employeeCount;
                _isActive = isActive;
                _prop1 = prop1;
                _prop2 = prop2;
                _prop3 = prop3;
                _prop4 = prop4;
                _prop5 = prop5;
                _prop6 = prop6;
                _prop7 = prop7;
                _contacts = contacts;
                _contracts = contracts;
                _products = products;
                _enumProp1 = enumProp1;
                _enumProp2 = enumProp2;
            }
            [ExcelExport(PropertyName = "Контакты")]
            public List<Contact> Contacts {
                get { return _contacts; }
            }
            [ExcelExport(PropertyName = "Сделки")]
            public List<Contract> Contracts {
                get { return _contracts; }
            }
            [ExcelExport(PropertyName = "Продукты")]
            public List<Product> Products {
                get { return _products; }
            }
            [ExcelExport(PropertyName = "Продукты")]
            public List<EnumProp1> EnumProps1 {
                get { return _enumProp1; }
            }
            [ExcelExport(PropertyName = "Продукты")]
            public List<EnumProp2> EnumProps2 {
                get { return _enumProp2; }
            }
            [ExcelExport(PropertyName = "Название")]
            public string Title {
                get { return _title; }
            }
            [ExcelExport(PropertyName = "Дата регистрации")]
            public DateTime RegistrationDate {
                get { return _registrationDate; }
            }
            [ExcelExport(PropertyName = "Телефон")]
            public string Phone {
                get { return _phone; }
            }
            [ExcelExport(PropertyName = "ОКАТО")]
            public string Okato {
                get { return _okato; }
            }
            [ExcelExport(PropertyName = "ИНН", IsExportable = false)]
            public string Inn {
                get { return _inn; }
            }
            [ExcelExport(PropertyName = "Прибыль за прошлый год")]
            public decimal Revenue {
                get { return _revenue; }
            }
            [ExcelExport(IsExportable = false)]
            public int EmployeeCount {
                get { return _employeeCount; }
            }
            [ExcelExport(PropertyName = "Удалена")]
            public bool IsActive {
                get { return _isActive; }
            }
            [ExcelExport(PropertyName = "Свойство 1")]
            public int Prop1 {
                get { return _prop1; }
            }
            [ExcelExport(PropertyName = "Свойство 2")]
            public string Prop2 {
                get { return _prop2; }
            }
            [ExcelExport(PropertyName = "Свойство 3")]
            public bool Prop3 {
                get { return _prop3; }
            }
            [ExcelExport(PropertyName = "Свойство 4")]
            public int Prop4 {
                get { return _prop4; }
            }
            [ExcelExport(PropertyName = "Свойство 5")]
            public string Prop5 {
                get { return _prop5; }
            }
            [ExcelExport(PropertyName = "Свойство 6")]
            public bool Prop6 {
                get { return _prop6; }
            }
            [ExcelExport(PropertyName = "Свойство 7")]
            public int Prop7 {
                get { return _prop7; }
            }

            [ExcelExportClassName(Name = "Контакт")]
            public sealed class Contact {
                private readonly string _title;
                private readonly string _email;
                public Contact(string title, string email) {
                    _title = title;
                    _email = email;
                }
                [ExcelExport(PropertyName = "Название")]
                public string Title {
                    get { return _title; }
                }
                [ExcelExport(PropertyName = "Email")]
                public string Email {
                    get { return _email; }
                }
            }
            [ExcelExportClassName(Name = "Сделка")]
            public sealed class Contract {
                private readonly DateTime _beginDate;
                private readonly DateTime _endDate;
                private readonly bool _status;
                public Contract(DateTime beginDate, DateTime endDate, bool status) {
                    _beginDate = beginDate;
                    _endDate = endDate;
                    _status = status;
                }
                [ExcelExport(PropertyName = "Дата начала")]
                public DateTime BeginDate {
                    get { return _beginDate; }
                }
                [ExcelExport(PropertyName = "Дата окончания")]
                public DateTime EndDate {
                    get { return _endDate; }
                }
                [ExcelExport(PropertyName = "Статус")]
                public bool Status {
                    get { return _status; }
                }
            }
            [ExcelExportClassName(Name = "Продукт")]
            public sealed class Product {
                private readonly string _title;
                private readonly int _amount;
                public Product(string title, int amount) {
                    _title = title;
                    _amount = amount;
                }
                [ExcelExport(PropertyName = "Наименование продукта")]
                public string Title {
                    get { return _title; }
                }
                [ExcelExport(PropertyName = "Количество")]
                public int Amount {
                    get { return _amount; }
                }
            }

            [ExcelExportClassName(Name = "Переч. свойство1")]
            public sealed class EnumProp1 {
                private readonly string _field1;
                private readonly int _field2;
                private readonly bool _field3;
                public EnumProp1(string field1, int field2, bool field3) {
                    _field1 = field1;
                    _field2 = field2;
                    _field3 = field3;
                }
                [ExcelExport(PropertyName = "Поле1")]
                public string Field1 {
                    get { return _field1; }
                }
                [ExcelExport(PropertyName = "Поле2")]
                public int Field2 {
                    get { return _field2; }
                }
                [ExcelExport(PropertyName = "Поле3")]
                public bool  Field3 {
                    get { return _field3; }
                }
            }

            [ExcelExportClassName(Name = "Перечисл. свойство2")]
            public sealed class EnumProp2 {
                private readonly string _field4;
                private readonly int _field5;
                private readonly bool _field6;
                private readonly string _field7;
                private readonly int _field8;
                public EnumProp2(string field4, int field5, bool field6, string field7, int field8) {
                    _field4 = field4;
                    _field5 = field5;
                    _field6 = field6;
                    _field7 = field7;
                    _field8 = field8;
                }
                [ExcelExport(PropertyName = "Поле4")]
                public string Field4 {
                    get { return _field4; }
                }
                [ExcelExport(PropertyName = "Поле5")]
                public int Field5 {
                    get { return _field5; }
                }
                [ExcelExport(PropertyName = "Поле6")]
                public bool  Field6 {
                    get { return _field6; }
                }
                [ExcelExport(PropertyName = "Поле7")]
                public string Field7 {
                    get { return _field7; }
                }
                [ExcelExport(PropertyName = "Поле8")]
                public int Field8 {
                    get { return _field8; }
                }
            }
        }
        public sealed class TestData {
            private readonly TableConfigurationBuilder<ClientExampleModel> _configurationBuilder;
            private readonly List<ClientExampleModel> _models;

            public TestData(TableConfigurationBuilder<ClientExampleModel> configurationBuilder, List<ClientExampleModel> models) {
                _configurationBuilder = configurationBuilder;
                _models = models;
            }

            public TableConfigurationBuilder<ClientExampleModel> ConfigurationBuilder {
                get { return _configurationBuilder; }
            }

            public List<ClientExampleModel> Models {
                get { return _models; }
            }
        }
    }
}