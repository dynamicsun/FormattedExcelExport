using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;


namespace FormattedExcelExport.Example {
	class Program {
		static void Main(string[] args) {
			var confBuilder = new TableConfigurationBuilder<ClientExampleModel>("Клиент", new CultureInfo("ru-RU"));

			confBuilder.RegisterColumnIf(true, "Название", x => x.Title);
			confBuilder.RegisterColumnIf(true, "Дата регистрации", x => x.RegistrationDate);
			confBuilder.RegisterColumnIf(true, "Телефон", x => x.Phone);

			var contact = confBuilder.RegisterChild("Контакт", x => x.Contacts);
			contact.RegisterColumnIf(true, "Название", x => x.Title);
			contact.RegisterColumnIf(true, "Email", x => x.Email);

			List<ClientExampleModel> models = InitializeModel();

			TableWriterStyle style = new TableWriterStyle();
			MemoryStream ms = TableWriterComplex.Write(new ExcelTableWriterComplex(style), models, confBuilder.Value);

			using (FileStream file = new FileStream("test.xls", FileMode.Create, System.IO.FileAccess.Write)) {
				byte[] bytes = new byte[ms.Length];
				ms.Read(bytes, 0, (int)ms.Length);
				file.Write(bytes, 0, bytes.Length);
				ms.Close();
			}
		}
		private static List<ClientExampleModel> InitializeModel() {
			return new List<ClientExampleModel> {
				new ClientExampleModel(
					"Первая компания", 
					DateTime.Now, 
					"+7 333 4442 00", 
					new List<ClientExampleModel.Contact> {
						new ClientExampleModel.Contact("Иван", "ivan@mail.ru")
					}),
				new ClientExampleModel(
					"Вторая компания", 
					DateTime.Now, 
					"+7 222 1124 44", 
					new List<ClientExampleModel.Contact> {
						new ClientExampleModel.Contact("Олег", "oleg@mail.ru")
					})		
			};
		}

		public class ClientExampleModel {
			private readonly string _title;
			private readonly DateTime _registrationDate;
			private readonly string _phone;
			private readonly List<Contact> _contacts;
			public ClientExampleModel(string title, DateTime registrationDate, string phone, List<Contact> contacts) {
				_title = title;
				_registrationDate = registrationDate;
				_phone = phone;
				_contacts = contacts;
			}

			public List<Contact> Contacts {
				get { return _contacts; }
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
		}
	}
}
