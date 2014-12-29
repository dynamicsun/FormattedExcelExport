using System.Collections.Generic;

namespace FormattedExcelExport.Tests.Models {
	internal sealed class Contact {
		public string FisrtName { get; set; }
		public string LastName { get; set; }

		public string Phone { get; set; }

		internal IReadOnlyList<Contact> EmptyList {
			get {
				var contacts = new List<Contact>();

				contacts.Add(new Contact {
					FisrtName = "Mike",
					LastName = "Smith",
					Phone = "+1 928 439-12-12"
				});

				contacts.Add(new Contact {
					FisrtName = "Alex",
					LastName = "Grachevsky",
					Phone = "+7 916 439-12-15 (ext. 284)"
				});

				contacts.Add(new Contact {
					FisrtName = "Marina",
					LastName = "Su",
					Phone = "+3 4321 19-20"
				});

				return contacts;
			}
		}

	}
}
