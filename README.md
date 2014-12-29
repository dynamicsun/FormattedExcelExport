.NET Formatted excel export (C#)
====================

![Warning](https://cdn0.iconfinder.com/data/icons/app_iconset_creative_nerds/16/warning.png) This is the public mirror of our internal repository.

***
###Table of contents###
1. [Description](#description). 
2. [Examples](#examples).

####Description##
The *formatted excel export* library may help you to export plain tables from *IEnumerable* objects to Excel files (both *.xls* and *.xlsx* formats). Faster then ever. 

![Tip](https://cdn2.iconfinder.com/data/icons/fugue/icon/light_bulb.png) Also you can export to a [DSV (CSV, TSV, etc.)](http://en.wikipedia.org/wiki/Delimiter-separated_values) file.

####Examples##
You want to export your contacts table.
You've got class *Contact* where export's fields are marked by *[ExcelExportAttribute]*:

	public sealed class Contact {
	    public int Id { get; set; }
	    
		[ExcelExport] public string FisrtName { get; set; }
		[ExcelExport] public string LastName { get; set; }

		[ExcelExport] public string Phone { get; set; }
	}

All that you need is to grab your contacts from anywhere your want, put them to *IEnumerable* and export them to Excel like this:

    var stream = ReflectionWriterSimple.Write(contacts, new XlsxTableWriterSimple(new TableWriterStyle()), new CultureInfo("en-US"));
 
That's all. After that you get a stream that represents an Excel file with default styles and "en-US" culture.

![Default export](http://i.imgur.com/kmFIyf7.png)

What if your class *Contact* was designed to have multiple phone numbers?
Your *Contact* class looks like:

	public sealed class Contact {
		[ExcelExport] public string FisrtName { get; set; }
		[ExcelExport] public string LastName { get; set; }

		public IReadOnlyList<Phone> Phones { get; set; }
	}
	
And alsow you've got *Phone* class now:

	public sealed class Phone {
		[ExcelExport(PropertyName = "Phone")] public string Number { get; set; }
	}
	
That's all. Now you can use the same method for export:

    var stream = ReflectionWriterSimple.Write(contacts, new XlsxTableWriterSimple(new TableWriterStyle()), new CultureInfo("en-US"));
    
After that all phones will be exported with suffix number: 'Phone1', 'Phone2', etc.

![Default export with enumeration](http://i.imgur.com/vo1wBML.png)

What if you want to change fields names? You can use *ProperyName* to change default column's name. Just see below:

	public sealed class Contact {
		[ExcelExport(PropertyName = "First name")] public string FisrtName { get; set; }
		[ExcelExport(PropertyName = "Last name")] public string LastName { get; set; }

		public IReadOnlyList<Phone> Phones { get; set; }
	}
	
Result:

![Property name export](http://i.imgur.com/qknQVYR.png)
