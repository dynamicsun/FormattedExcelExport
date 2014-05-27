﻿using System.Linq;
using NPOI.HSSF.Record;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using FormattedExcelExport.Configuaration;
using FormattedExcelExport.Reflection;
using FormattedExcelExport.Style;
using FormattedExcelExport.TableWriters;

namespace FormattedExcelExport.Tests {
	[TestFixture]
	public class Test {
        /*static void Main(string[] args) {
            var confBuilder = new TableConfigurationBuilder<ClientExampleModel>("Клиент", new CultureInfo("ru-RU"));
            TableWriterStyle condStyle = new TableWriterStyle();
            condStyle.RegularCell.BackgroundColor = new AdHocCellStyle.Color(255, 0, 0);
            TableWriterStyle condStyle2 = new TableWriterStyle();
            condStyle2.RegularCell.BackgroundColor = new AdHocCellStyle.Color(0, 255, 0);
            TableWriterStyle condStyle3 = new TableWriterStyle();
            condStyle3.RegularChildCell.BackgroundColor = new AdHocCellStyle.Color(0, 0, 255);

            confBuilder.RegisterColumn("Название", x => x.Title, new TableConfigurationBuilder<ClientExampleModel>.ConditionTheme(condStyle, x => x.Title == "Вторая компания"));
            confBuilder.RegisterColumn("Дата регистрации", x => x.RegistrationDate);
            confBuilder.RegisterColumn("Телефон", x => x.Phone, new TableConfigurationBuilder<ClientExampleModel>.ConditionTheme(condStyle2, x => x.Okato == "OPEEHBSSDD"));
            confBuilder.RegisterColumn("ИНН", x => x.Inn);
            confBuilder.RegisterColumn("Окато", x => x.Okato);

            var contact = confBuilder.RegisterChild("Контакт", x => x.Contacts);
            contact.RegisterColumn("Название", x => x.Title, new TableConfigurationBuilder<ClientExampleModel.Contact>.ConditionTheme(condStyle3, x => x.Title.StartsWith("О")));
            contact.RegisterColumn("Email", x => x.Email);

            var contract = confBuilder.RegisterChild("Контракт", x => x.Contracts);
            contract.RegisterColumn("Дата начала", x => x.BeginDate);
            contract.RegisterColumn("Дата окончания", x => x.EndDate);
            contract.RegisterColumn("Статус", x => x.Status, new TableConfigurationBuilder<ClientExampleModel.Contract>.ConditionTheme(new TableWriterStyle(), x => true));

            var product = confBuilder.RegisterChild("Продукт", x => x.Products);
            product.RegisterColumn("Наименование", x => x.Title);
            product.RegisterColumn("Количество", x => x.Amount);

            List<ClientExampleModel> models = InitializeModels();

            MemoryStream ms = TableWriterComplex.Write(new DsvTableWriterComplex(), models, confBuilder.Value);
            WriteToFile(ms, "TestComplex.txt");

            TableWriterStyle style = new TableWriterStyle();
            ms = TableWriterComplex.Write(new XlsTableWriterComplex(style), models, confBuilder.Value);
            WriteToFile(ms, "TestComplex.xls");

            ms = TableWriterSimple.Write(new DsvTableWriterSimple(), models, confBuilder.Value);
            WriteToFile(ms, "TestSimple.txt");

            ms = TableWriterSimple.Write(new XlsTableWriterSimple(style), models, confBuilder.Value);
            WriteToFile(ms, "TestSimple.xls");

            TableWriterSimple.Write(new XlsxTableWriterSimple(style), models, confBuilder.Value);
            TableWriterComplex.Write(new XlsxTableWriterComplex(style), models, confBuilder.Value);

            ms = ReflectionWriterSimple.Write(models, new DsvTableWriterSimple(), new CultureInfo("ru-Ru"));
            WriteToFile(ms, "TestReflectionSimple.txt");

            ms = ReflectionWriterSimple.Write(models, new XlsTableWriterSimple(new TableWriterStyle()), new CultureInfo("ru-Ru"));
            WriteToFile(ms, "TestReflectionSimple.xls");

            ms = ReflectionWriterComplex.Write(models, new DsvTableWriterComplex(), new CultureInfo("ru-Ru"));
            WriteToFile(ms, "TestReflectionSimple.txt");

            ms = ReflectionWriterComplex.Write(models, new XlsTableWriterComplex(new TableWriterStyle()), new CultureInfo("ru-Ru"));
            WriteToFile(ms, "TestReflectionComplex.xls");
        }*/

        private static List<ClientExampleModel> CreateTestExample(out MemoryStream ms) {
            TableConfigurationBuilder<ClientExampleModel> confBuilder =
                new TableConfigurationBuilder<ClientExampleModel>("Клиент", new CultureInfo("ru-RU"));
            TableWriterStyle condStyle = new TableWriterStyle();
            condStyle.RegularCell.BackgroundColor = new AdHocCellStyle.Color(255, 0, 0);
            TableWriterStyle condStyle2 = new TableWriterStyle();
            condStyle2.RegularCell.BackgroundColor = new AdHocCellStyle.Color(0, 255, 0);
            TableWriterStyle condStyle3 = new TableWriterStyle();
            condStyle3.RegularChildCell.BackgroundColor = new AdHocCellStyle.Color(0, 0, 255);

            confBuilder.RegisterColumn("Название", x => x.Title,
                new TableConfigurationBuilder<ClientExampleModel>.ConditionTheme(condStyle, x => x.Title == "Вторая компания"));
            confBuilder.RegisterColumn("Дата регистрации", x => x.RegistrationDate);
            confBuilder.RegisterColumn("Телефон", x => x.Phone,
                new TableConfigurationBuilder<ClientExampleModel>.ConditionTheme(condStyle2, x => x.Okato == "OPEEHBSSDD"));
            confBuilder.RegisterColumn("ИНН", x => x.Inn);
            confBuilder.RegisterColumn("Окато", x => x.Okato);

            TableConfigurationBuilder<ClientExampleModel.Contact> contact = confBuilder.RegisterChild("Контакт", x => x.Contacts);
            contact.RegisterColumn("Название", x => x.Title,
                new TableConfigurationBuilder<ClientExampleModel.Contact>.ConditionTheme(condStyle3,
                    x => x.Title.StartsWith("О")));
            contact.RegisterColumn("Email", x => x.Email);

            TableConfigurationBuilder<ClientExampleModel.Contract> contract = confBuilder.RegisterChild("Контракт",
                x => x.Contracts);
            contract.RegisterColumn("Дата начала", x => x.BeginDate);
            contract.RegisterColumn("Дата окончания", x => x.EndDate);
            contract.RegisterColumn("Статус", x => x.Status,
                new TableConfigurationBuilder<ClientExampleModel.Contract>.ConditionTheme(new TableWriterStyle(), x => true));

            TableConfigurationBuilder<ClientExampleModel.Product> product = confBuilder.RegisterChild("Продукт", x => x.Products);
            product.RegisterColumn("Наименование", x => x.Title);
            product.RegisterColumn("Количество", x => x.Amount);
            List<ClientExampleModel> models = InitializeModels();

            TableWriterStyle style = new TableWriterStyle();

            ms = TableWriterSimple.Write(new XlsTableWriterSimple(style), models, confBuilder.Value);
            return models;
        }

		[Test]
		public void ExcelSimpleExport() {
		    MemoryStream ms;
		    var models = CreateTestExample(out ms);
		    WriteToFile(ms, "TestSimple.xls");
            
		    HSSFWorkbook xlsFile;
            using (FileStream file = new FileStream("TestSimple.xls", FileMode.Open, FileAccess.Read)) {
		        xlsFile = new HSSFWorkbook(file);
		    }

		    ISheet sheet = xlsFile.GetSheetAt(0);
            for (int rowNumber = 1; rowNumber <= sheet.LastRowNum; rowNumber++) {
                IRow row = sheet.GetRow(rowNumber);

                int listNumber = rowNumber - 1;
                int countContacts = (from ClientExampleModel model in models
                                     select model.Contacts.Count).Max();
                int countContracts = (from ClientExampleModel model in models
                                      select model.Contracts.Count).Max();
                int countProducts = (from ClientExampleModel model in models
                                     select model.Products.Count).Max();
                Assert.AreNotEqual(row, null);
                Assert.AreEqual(row.GetCell(0).StringCellValue, models[listNumber].Title);
                Assert.AreEqual(row.GetCell(1).StringCellValue, models[listNumber].RegistrationDate.ToString("d MMMMM yyyy ", new CultureInfo("ru-RU")) + "г.");
                Assert.AreEqual(row.GetCell(2).StringCellValue, models[listNumber].Phone);
                Assert.AreEqual(row.GetCell(3).StringCellValue, models[listNumber].Inn);
                Assert.AreEqual(row.GetCell(4).StringCellValue, models[listNumber].Okato);

                for (int contactNumber = 0; contactNumber < models[listNumber].Contacts.Count; contactNumber++) {
                    Assert.AreEqual(row.GetCell(5 + (contactNumber * 2)).StringCellValue, models[listNumber].Contacts[contactNumber].Title);
                    Assert.AreEqual(row.GetCell(5 + (contactNumber * 2 + 1)).StringCellValue, models[listNumber].Contacts[contactNumber].Email);
                }

                for (int contractNumber = 0; contractNumber < models[listNumber].Contracts.Count; contractNumber++) {
                    Assert.AreEqual(row.GetCell(5 + (countContacts * 2) + (contractNumber * 3)).StringCellValue,
                        models[listNumber].Contracts[contractNumber].BeginDate.ToString("d MMMMM yyyy ", new CultureInfo("ru-RU")) + "г.");
                    Assert.AreEqual(row.GetCell(5 + (countContacts * 2) + (contractNumber * 3 + 1)).StringCellValue,
                        models[listNumber].Contracts[contractNumber].EndDate.ToString("d MMMMM yyyy ", new CultureInfo("ru-RU")) + "г.");
                    if (models[listNumber].Contracts[contractNumber].Status) {
                        Assert.AreEqual(
                            row.GetCell(5 + (countContacts * 2) + (contractNumber * 3 + 2)).StringCellValue,
                            "Да");
                    } else {
                        Assert.AreEqual(
                            row.GetCell(5 + (countContacts * 2) + (contractNumber * 3 + 2)).StringCellValue,
                            "Нет");
                    }
                }

                for (int productNumber = 0; productNumber < models[listNumber].Products.Count; productNumber++) {
                    Assert.AreEqual(row.GetCell(5 + (countContacts * 2) + (countContracts * 3) + (productNumber * 2)).StringCellValue,
                        models[listNumber].Products[productNumber].Title);
                    Assert.AreEqual(row.GetCell(5 + (countContacts * 2) + (countContracts * 3) + (productNumber * 2 + 1)).StringCellValue,
                        models[listNumber].Products[productNumber].Amount.ToString());
                }
            }
		}
        /*
	    [Test]
        public void TxtSimpleExport() {
            TableConfigurationBuilder<ClientExampleModel> confBuilder = new TableConfigurationBuilder<ClientExampleModel>("Клиент", new CultureInfo("ru-RU"));

            TableWriterStyle condStyle = new TableWriterStyle();
            condStyle.RegularCell.BackgroundColor = new AdHocCellStyle.Color(255, 0, 0);
            TableWriterStyle condStyle2 = new TableWriterStyle();
            condStyle2.RegularCell.BackgroundColor = new AdHocCellStyle.Color(0, 255, 0);
            TableWriterStyle condStyle3 = new TableWriterStyle();
            condStyle3.RegularChildCell.BackgroundColor = new AdHocCellStyle.Color(0, 0, 255);

            confBuilder.RegisterColumn("Название", x => x.Title, new TableConfigurationBuilder<ClientExampleModel>.ConditionTheme(condStyle, x => x.Title == "Вторая компания"));
            confBuilder.RegisterColumn("Дата регистрации", x => x.RegistrationDate);
            confBuilder.RegisterColumn("Телефон", x => x.Phone, new TableConfigurationBuilder<ClientExampleModel>.ConditionTheme(condStyle2, x => x.Okato == "OPEEHBSSDD"));
            confBuilder.RegisterColumn("ИНН", x => x.Inn);
            confBuilder.RegisterColumn("Окато", x => x.Okato);

            TableConfigurationBuilder<ClientExampleModel.Contact> contact = confBuilder.RegisterChild("Контакт", x => x.Contacts);
            contact.RegisterColumn("Название", x => x.Title, new TableConfigurationBuilder<ClientExampleModel.Contact>.ConditionTheme(condStyle3, x => x.Title.StartsWith("О")));
            contact.RegisterColumn("Email", x => x.Email);

            TableConfigurationBuilder<ClientExampleModel.Contract> contract = confBuilder.RegisterChild("Контракт", x => x.Contracts);
            contract.RegisterColumn("Дата начала", x => x.BeginDate);
            contract.RegisterColumn("Дата окончания", x => x.EndDate);
            contract.RegisterColumn("Статус", x => x.Status, new TableConfigurationBuilder<ClientExampleModel.Contract>.ConditionTheme(new TableWriterStyle(), x => true));

            TableConfigurationBuilder<ClientExampleModel.Product> product = confBuilder.RegisterChild("Продукт", x => x.Products);
            product.RegisterColumn("Наименование", x => x.Title);
            product.RegisterColumn("Количество", x => x.Amount);

            List<ClientExampleModel> models = InitializeModels();

            MemoryStream ms = TableWriterSimple.Write(new DsvTableWriterSimple(), models, confBuilder.Value);
            WriteToFile(ms, "TestSimple.txt");
        }

        protected ICellStyle ConvertToNpoiStyle(AdHocCellStyle adHocCellStyle, HSSFWorkbook workbook) {
            IFont cellFont = workbook.CreateFont();

            cellFont.FontName = adHocCellStyle.FontName;
            cellFont.FontHeightInPoints = adHocCellStyle.FontHeightInPoints;
            cellFont.IsItalic = adHocCellStyle.Italic;
            cellFont.Underline = adHocCellStyle.Underline ? FontUnderlineType.Single : FontUnderlineType.None;
            cellFont.Boldweight = (short)adHocCellStyle.BoldWeight;

            HSSFPalette palette = workbook.GetCustomPalette();
            HSSFColor similarColor = palette.FindSimilarColor(adHocCellStyle.FontColor.Red, adHocCellStyle.FontColor.Green, adHocCellStyle.FontColor.Blue);
            cellFont.Color = similarColor.GetIndex();

            ICellStyle cellStyle = workbook.CreateCellStyle();
            cellStyle.SetFont(cellFont);

            if (adHocCellStyle.BackgroundColor != null) {
                similarColor = palette.FindSimilarColor(adHocCellStyle.BackgroundColor.Red, adHocCellStyle.BackgroundColor.Green, adHocCellStyle.BackgroundColor.Blue);
                cellStyle.FillForegroundColor = similarColor.GetIndex();
                cellStyle.FillPattern = FillPattern.SolidForeground;
            }
            return cellStyle;
        }*/

        private static void WriteToFile(MemoryStream ms, string fileName) {
            using (FileStream file = new FileStream(fileName, FileMode.Create, FileAccess.Write)) {
                byte[] bytes = new byte[ms.Length];
                ms.Read(bytes, 0, (int)ms.Length);
                file.Write(bytes, 0, bytes.Length);
                ms.Close();
            }
        }
        private static List<ClientExampleModel> InitializeModels() {
            return new List<ClientExampleModel> {
				new ClientExampleModel(
					"Первая компания", 
					DateTime.Now, 
					"+7 333 4442 00", 
					"9040043234562",
					"OPEEHBSSDD",
					2352666,
					336,
					true,
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
					new List<ClientExampleModel.Contact> {
						new ClientExampleModel.Contact("Олег", "oleg@mail.ru"),
						new ClientExampleModel.Contact("Анна", ""),
						new ClientExampleModel.Contact("Николай", "nikolay@mail.ru")
					},
					new List<ClientExampleModel.Contract> {
						new ClientExampleModel.Contract(new DateTime(1999, 1, 7), new DateTime(2007, 3, 22), true),
						new ClientExampleModel.Contract(new DateTime(1989, 2, 4), new DateTime(2012, 11, 20), false)
					},
					new List<ClientExampleModel.Product> ()),
		        new ClientExampleModel(
					"Первая компания", 
					DateTime.Now, 
					"+7 333 4442 00", 
					"9040043234562",
					"OPEEHBSSDD",
					2352666,
					336,
					true,
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
						new ClientExampleModel.Product("Лук", 100),
                        new ClientExampleModel.Product("Курица", 250)
					}),
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
        }
	}
}
