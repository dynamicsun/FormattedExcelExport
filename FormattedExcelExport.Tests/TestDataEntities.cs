using System;
using System.Collections.Generic;
using System.Globalization;
using FormattedExcelExport.Configuaration;
using FormattedExcelExport.Reflection;
using FormattedExcelExport.Style;

namespace FormattedExcelExport.Tests {
    internal static class TestDataEntities {
        internal static TestData CreateSimpleTestData(bool style = false) {
            TableConfigurationBuilder<ClientExampleModel> dataStructure = CreateSimpleTestDataConfigurationBuilder(style);           
            var data = CreateSimpleTestDataModels();
            return new TestData(dataStructure, data);
        }
        private static TableConfigurationBuilder<ClientExampleModel> CreateSimpleTestDataConfigurationBuilder(bool style = false) {
            if (style) {
                TableWriterStyle condStyle = new TableWriterStyle();
                condStyle.RegularCell.BackgroundColor = new AdHocCellStyle.Color(255, 0, 0);
                TableWriterStyle condStyle2 = new TableWriterStyle();
                condStyle2.RegularCell.BackgroundColor = new AdHocCellStyle.Color(0, 255, 0);
                TableWriterStyle condStyle3 = new TableWriterStyle();
                condStyle3.RegularChildCell.BackgroundColor = new AdHocCellStyle.Color(0, 0, 255);
                
                TableConfigurationBuilder<ClientExampleModel> confBuilder = new TableConfigurationBuilder<ClientExampleModel>("������", new CultureInfo("ru-RU"));
                confBuilder.RegisterColumn("��������", x => x.Title, new TableConfigurationBuilder<ClientExampleModel>.ConditionTheme(condStyle, x => x.Title == "������ ��������"));
                confBuilder.RegisterColumn("���� �����������", x => x.RegistrationDate);
                confBuilder.RegisterColumn("�������", x => x.Phone, new TableConfigurationBuilder<ClientExampleModel>.ConditionTheme(condStyle2, x => x.Okato == "OPEEHBSSDD"));
                confBuilder.RegisterColumn("���", x => x.Inn);
                confBuilder.RegisterColumn("�����", x => x.Okato);

                TableConfigurationBuilder<ClientExampleModel.Contact> contact = confBuilder.RegisterChild("�������", x => x.Contacts);
                contact.RegisterColumn("��������", x => x.Title, new TableConfigurationBuilder<ClientExampleModel.Contact>.ConditionTheme(condStyle3, x => x.Title.StartsWith("�")));
                contact.RegisterColumn("Email", x => x.Email);

                TableConfigurationBuilder<ClientExampleModel.Contract> contract = confBuilder.RegisterChild("��������", x => x.Contracts);
                contract.RegisterColumn("���� ������", x => x.BeginDate);
                contract.RegisterColumn("���� ���������", x => x.EndDate);
                contract.RegisterColumn("������", x => x.Status, new TableConfigurationBuilder<ClientExampleModel.Contract>.ConditionTheme(new TableWriterStyle(), x => true));

                TableConfigurationBuilder<ClientExampleModel.Product> product = confBuilder.RegisterChild("�������", x => x.Products);
                product.RegisterColumn("������������", x => x.Title);
                product.RegisterColumn("����������", x => x.Amount);
                return confBuilder;
            }
            else {
                TableConfigurationBuilder<ClientExampleModel> confBuilder = new TableConfigurationBuilder<ClientExampleModel>("������", new CultureInfo("ru-RU"));
                confBuilder.RegisterColumn("��������", x => x.Title);
                confBuilder.RegisterColumn("���� �����������", x => x.RegistrationDate);
                confBuilder.RegisterColumn("�������", x => x.Phone);
                confBuilder.RegisterColumn("���", x => x.Inn);
                confBuilder.RegisterColumn("�����", x => x.Okato);

                TableConfigurationBuilder<ClientExampleModel.Contact> contact = confBuilder.RegisterChild("�������", x => x.Contacts);
                contact.RegisterColumn("��������", x => x.Title);
                contact.RegisterColumn("Email", x => x.Email);

                TableConfigurationBuilder<ClientExampleModel.Contract> contract = confBuilder.RegisterChild("��������", x => x.Contracts);
                contract.RegisterColumn("���� ������", x => x.BeginDate);
                contract.RegisterColumn("���� ���������", x => x.EndDate);
                contract.RegisterColumn("������", x => x.Status);

                TableConfigurationBuilder<ClientExampleModel.Product> product = confBuilder.RegisterChild("�������", x => x.Products);
                product.RegisterColumn("������������", x => x.Title);
                product.RegisterColumn("����������", x => x.Amount);
                return confBuilder;
            }
        }
        internal static List<ClientExampleModel> CreateSimpleTestDataModels() {
            return new List<ClientExampleModel> {
				new ClientExampleModel(
					"������ ��������", 
					DateTime.Now, 
					"+7 333 4442 00", 
					"9040043234562",
					"OPEEHBSSDD",
					2352666,
					336,
					true,
					new List<ClientExampleModel.Contact> {
						new ClientExampleModel.Contact("�����", "olga@mail.ru"),
						new ClientExampleModel.Contact("����", "ivan@mail.ru")
					},
					new List<ClientExampleModel.Contract> {
						new ClientExampleModel.Contract(new DateTime(1999, 1, 7), new DateTime(2009, 11, 9), false),
						new ClientExampleModel.Contract(new DateTime(1989, 2, 4), DateTime.Now, true)
					},
					new List<ClientExampleModel.Product> {
						new ClientExampleModel.Product("���������", 20),
						new ClientExampleModel.Product("���", 100)
					}),
				new ClientExampleModel(
					"������ ��������", 
					DateTime.Now, 
					"+7 222 1124 44", 
					"5953043385461",
					"JsKSLPKKHSS",
					599988,
					59,
					false,
					new List<ClientExampleModel.Contact> {
						new ClientExampleModel.Contact("����", "oleg@mail.ru"),
						new ClientExampleModel.Contact("����", ""),
						new ClientExampleModel.Contact("�������", "nikolay@mail.ru")
					},
					new List<ClientExampleModel.Contract> {
						new ClientExampleModel.Contract(new DateTime(1999, 1, 7), new DateTime(2007, 3, 22), true),
						new ClientExampleModel.Contract(new DateTime(1989, 2, 4), new DateTime(2012, 11, 20), false)
					},
					new List<ClientExampleModel.Product> ())
			};
        }
        [ExcelExportClassName(Name = "������")]
        public class ClientExampleModel {
            private readonly string _title;
            private readonly DateTime _registrationDate;
            private readonly string _phone;
            private readonly string _inn;
            private readonly string _okato;
            private readonly decimal _revenue;
            private readonly int _employeeCount;
            private readonly bool _isActive;
            private readonly List<Contact> _contacts;
            private readonly List<Contract> _contracts;
            private readonly List<Product> _products;

            public ClientExampleModel(string title, DateTime registrationDate, string phone, string inn, string okato, decimal revenue, int employeeCount, bool isActive, List<Contact> contacts, List<Contract> contracts, List<Product> products) {
                _title = title;
                _registrationDate = registrationDate;
                _phone = phone;
                _inn = inn;
                _okato = okato;
                _revenue = revenue;
                _employeeCount = employeeCount;
                _isActive = isActive;
                _contacts = contacts;
                _contracts = contracts;
                _products = products;
            }
            [ExcelExport(PropertyName = "��������")]
            public List<Contact> Contacts {
                get { return _contacts; }
            }
            [ExcelExport(PropertyName = "������")]
            public List<Contract> Contracts {
                get { return _contracts; }
            }
            [ExcelExport(PropertyName = "��������")]
            public List<Product> Products {
                get { return _products; }
            }
            [ExcelExport(PropertyName = "��������")]
            public string Title {
                get { return _title; }
            }
            [ExcelExport(PropertyName = "���� �����������")]
            public DateTime RegistrationDate {
                get { return _registrationDate; }
            }
            [ExcelExport(PropertyName = "�������")]
            public string Phone {
                get { return _phone; }
            }
            [ExcelExport(PropertyName = "�����")]
            public string Okato {
                get { return _okato; }
            }
            [ExcelExport(PropertyName = "���", IsExportable = false)]
            public string Inn {
                get { return _inn; }
            }
            [ExcelExport(PropertyName = "������� �� ������� ���")]
            public decimal Revenue {
                get { return _revenue; }
            }
            [ExcelExport(IsExportable = false)]
            public int EmployeeCount {
                get { return _employeeCount; }
            }
            [ExcelExport(PropertyName = "�������")]
            public bool IsActive {
                get { return _isActive; }
            }

            [ExcelExportClassName(Name = "�������")]
            public sealed class Contact {
                private readonly string _title;
                private readonly string _email;
                public Contact(string title, string email) {
                    _title = title;
                    _email = email;
                }
                [ExcelExport(PropertyName = "��������")]
                public string Title {
                    get { return _title; }
                }
                [ExcelExport(PropertyName = "Email")]
                public string Email {
                    get { return _email; }
                }
            }
            [ExcelExportClassName(Name = "������")]
            public sealed class Contract {
                private readonly DateTime _beginDate;
                private readonly DateTime _endDate;
                private readonly bool _status;
                public Contract(DateTime beginDate, DateTime endDate, bool status) {
                    _beginDate = beginDate;
                    _endDate = endDate;
                    _status = status;
                }
                [ExcelExport(PropertyName = "���� ������")]
                public DateTime BeginDate {
                    get { return _beginDate; }
                }
                [ExcelExport(PropertyName = "���� ���������")]
                public DateTime EndDate {
                    get { return _endDate; }
                }
                [ExcelExport(PropertyName = "������")]
                public bool Status {
                    get { return _status; }
                }
            }
            [ExcelExportClassName(Name = "�������")]
            public sealed class Product {
                private readonly string _title;
                private readonly int _amount;
                public Product(string title, int amount) {
                    _title = title;
                    _amount = amount;
                }
                [ExcelExport(PropertyName = "������������ ��������")]
                public string Title {
                    get { return _title; }
                }
                [ExcelExport(PropertyName = "����������")]
                public int Amount {
                    get { return _amount; }
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