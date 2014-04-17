using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;


namespace FormattedExcelExport.Example {
	class Program {
		static void Main(string[] args) {
			var confBuilder = new TableConfigurationBuilder<ClientExampleModel>("Клиент", new CultureInfo("ru-RU"));
			TableWriterStyle condStyle = new TableWriterStyle();
			condStyle.RegularCell.BackgroundColor = new StyleSettings.Color(255, 0, 0);
			TableWriterStyle condStyle2 = new TableWriterStyle();
			condStyle2.RegularCell.BackgroundColor = new StyleSettings.Color(0, 255, 0);
			TableWriterStyle condStyle3 = new TableWriterStyle();
			condStyle3.RegularChildCell.BackgroundColor = new StyleSettings.Color(0, 0, 255);

			confBuilder.RegisterColumnIf(true, "Название", x => x.Title, new TableConfigurationBuilder<ClientExampleModel>.ConditionTheme(condStyle, x => x.Title == "Вторая компания"));
			confBuilder.RegisterColumnIf(true, "Дата регистрации", x => x.RegistrationDate);
			confBuilder.RegisterColumnIf(true, "Телефон", x => x.Phone, new TableConfigurationBuilder<ClientExampleModel>.ConditionTheme(condStyle2, x => x.Okato == "OPEEHBSSDD"));
			confBuilder.RegisterColumnIf(true, "ИНН", x => x.Inn);
			confBuilder.RegisterColumnIf(true, "Окато", x => x.Okato);

			var contact = confBuilder.RegisterChild("Контакт", x => x.Contacts);
			contact.RegisterColumnIf(true, "Название", x => x.Title, new TableConfigurationBuilder<ClientExampleModel.Contact>.ConditionTheme(condStyle3, x => x.Title.StartsWith("О")));
			contact.RegisterColumnIf(true, "Email", x => x.Email);

			var contract = confBuilder.RegisterChild("Контракт", x => x.Contracts);
			contract.RegisterColumnIf(true, "Дата начала", x => x.BeginDate);
			contract.RegisterColumnIf(true, "Дата окончания", x => x.EndDate);
			contract.RegisterColumnIf(true, "Статус", x => x.Status);

			var product = confBuilder.RegisterChild("Продукт", x => x.Products);
			product.RegisterColumnIf(true, "Наименование", x => x.Title);
			product.RegisterColumnIf(true, "Количество", x => x.Amount);

			List<ClientExampleModel> testModels = InitializeModels();

			MemoryStream ms = TableWriterComplex.Write(new CsvTableWriterComplex(), testModels, confBuilder.Value);
			WriteToFile(ms, "test.txt");

			TableWriterStyle style = new TableWriterStyle();
			ms = TableWriterComplex.Write(new ExcelTableWriterComplex(style), testModels, confBuilder.Value);
			WriteToFile(ms, "test.xls");
		}

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
					"9040043234562",
					"JsKSLPKKHSS",
					new List<ClientExampleModel.Contact> {
						new ClientExampleModel.Contact("Олег", "oleg@mail.ru"),
						new ClientExampleModel.Contact("Анна", ""),
						new ClientExampleModel.Contact("Николай", "nikolay@mail.ru")
					},
					new List<ClientExampleModel.Contract> {
						new ClientExampleModel.Contract(new DateTime(1999, 1, 7), new DateTime(2007, 3, 22), true),
						new ClientExampleModel.Contract(new DateTime(1989, 2, 4), new DateTime(2012, 11, 20), false)
					},
					new List<ClientExampleModel.Product> ())		
			};
		}

		public class ClientExampleModel {
			private readonly string _title;
			private readonly DateTime _registrationDate;
			private readonly string _phone;
			private readonly string _inn;
			private readonly string _okato;
			private readonly List<Contact> _contacts;
			private readonly List<Contract> _contracts;
			private readonly List<Product> _products;
			public ClientExampleModel(string title, DateTime registrationDate, string phone, string inn, string okato, List<Contact> contacts, List<Contract> contracts, List<Product> products) {
				_title = title;
				_registrationDate = registrationDate;
				_phone = phone;
				_inn = inn;
				_okato = okato;
				_contacts = contacts;
				_contracts = contracts;
				_products = products;
			}

			public List<Contact> Contacts {
				get { return _contacts; }
			}
			public List<Contract> Contracts {
				get { return _contracts; }
			}
			public string Title {
				get { return _title; }
			}
			public DateTime RegistrationDate {
				get { return _registrationDate; }
			}

			public string Phone {
				get { return _phone; }
			}
			public string Okato {
				get { return _okato; }
			}
			public string Inn {
				get { return _inn; }
			}
			public List<Product> Products {
				get { return _products; }
			}
			public sealed class Contact {
				private readonly string _title;
				private readonly string _email;
				public Contact(string title, string email) {
					_title = title;
					_email = email;
				}
				public string Title {
					get { return _title; }
				}
				public string Email {
					get { return _email; }
				}
			}
			public sealed class Contract {
				private readonly DateTime _beginDate;
				private readonly DateTime _endDate;
				private readonly bool _status;
				public Contract(DateTime beginDate, DateTime endDate, bool status) {
					_beginDate = beginDate;
					_endDate = endDate;
					_status = status;
				}
				public DateTime BeginDate {
					get { return _beginDate; }
				}
				public DateTime EndDate {
					get { return _endDate; }
				}
				public bool Status {
					get { return _status; }
				}
			}
			public sealed class Product {
				private readonly string _title;
				private readonly int _amount;
				public Product(string title, int amount) {
					_title = title;
					_amount = amount;
				}
				public string Title {
					get { return _title; }
				}
				public int Amount {
					get { return _amount; }
				}
			}
		}
	}
}
