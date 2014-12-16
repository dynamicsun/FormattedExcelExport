using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using Extensions;
using FormattedExcelExport.Configuaration;
using FormattedExcelExport.Reflection;
using FormattedExcelExport.Style;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using System.IO;
using FormattedExcelExport.TableWriters;

namespace FormattedExcelExport.Tests {
	[TestFixture]
	public class Test {
	    public class FloatExpotClass {
	        [ExcelExport(PropertyName = "Имя значения")]
            public string Name { get; set; }
            [ExcelExport(PropertyName = "Само значение")]
            public float? Value { get; set; }
	    }

	    public class NullableTypes {
            [ExcelExport(PropertyName = "Int")]
            public int? IntValue { get; set; }
            [ExcelExport(PropertyName = "Float")]
            public float? FloatValue { get; set; }
            [ExcelExport(PropertyName = "Decimal")]
            public decimal? DecimalValue { get; set; }
            [ExcelExport(PropertyName = "Date")]
            public DateTime? DateValue { get; set; }
	    }
	    [Test]
	    public void ExportNullableNotReflectionTest() {
	        var nullTypeList = new List<NullableTypes>();
            nullTypeList.Add(new NullableTypes {
                DateValue = null,
                DecimalValue = (decimal?) 1.4544,
                FloatValue = (float?) 0.323232,
                IntValue = null
             });
            nullTypeList.Add(new NullableTypes {
                DateValue = new DateTime(2012, 11, 21),
                DecimalValue = null,
                FloatValue = null,
                IntValue = 12
            });
            var confBuilder = new TableConfigurationBuilder<NullableTypes>("Nullable", new CultureInfo("ru-RU"));
            confBuilder.RegisterColumn("Date", x => x.DateValue);
            confBuilder.RegisterColumn("Decimal", x => x.DecimalValue);
            confBuilder.RegisterColumn("Int", x => x.IntValue);
            confBuilder.RegisterColumn("Float", x => x.FloatValue);

            var filename = "TestNullableNotReflection.xls";
            DeleteTestFile(filename);
            var memoryStream = TableWriterSimple.Write(new XlsTableWriterSimple(), nullTypeList, confBuilder.Value);
            WriteToFile(memoryStream, filename);

	        var filenameXlsx = "TestNullableNotReflection.xlsx";
            DeleteTestFile(filenameXlsx);
            var memoryStreamXlsx = TableWriterSimple.Write(new XlsxTableWriterSimple(new TableWriterStyle()), nullTypeList, confBuilder.Value);
            WriteToFile(memoryStreamXlsx, filenameXlsx);

            var filenameXlsReflection = "TestNullableReflection.xls";
            DeleteTestFile(filenameXlsReflection);
            var memoryStream1 = ReflectionWriterSimple.Write(nullTypeList, new XlsTableWriterSimple(new TableWriterStyle()), new CultureInfo("ru-Ru"));
            WriteToFile(memoryStream1, filenameXlsReflection);

            var filenameXlsxReflection = "TestNullableReflection.xlsx";
            DeleteTestFile(filenameXlsxReflection);
            var memoryStream2 = ReflectionWriterSimple.Write(nullTypeList, new XlsxTableWriterSimple(new TableWriterStyle()), new CultureInfo("ru-Ru"));
            WriteToFile(memoryStream2, filenameXlsxReflection);

            var filenameXlsReflectionComplex = "TestNullableReflectionComplex.xls";
            DeleteTestFile(filenameXlsReflectionComplex);
            var memoryStreamComplex1 = ReflectionWriterComplex.Write(nullTypeList, new XlsTableWriterComplex(new TableWriterStyle()), new CultureInfo("ru-Ru"));
            WriteToFile(memoryStreamComplex1, filenameXlsReflectionComplex);

            var filenameXlsxReflectionComplex = "TestNullableReflectionComplex.xlsx";
            DeleteTestFile(filenameXlsxReflectionComplex);
            var memoryStreamComplex2 = ReflectionWriterComplex.Write(nullTypeList, new XlsxTableWriterComplex(new TableWriterStyle()), new CultureInfo("ru-Ru"));
            WriteToFile(memoryStreamComplex2, filenameXlsxReflectionComplex);
	    }
	    [Test]
	    public void ExportFloatFormatInReflection() {
	        var floatList = new List<FloatExpotClass> {
	            new FloatExpotClass {
	                Name = string.Empty,
	                Value = (float?) 2.1
	            },
	            new FloatExpotClass {
	                Name = "Значение два",
	                Value = (float?) 4.1,
	            },
	            new FloatExpotClass {
	                Name = "Значение три",
	                Value = null
	            }
	        };
	        const string filename = "FloatExport.xlsx";
            var memoryStream1 = ReflectionWriterSimple.Write(floatList, new XlsxTableWriterSimple(new TableWriterStyle()), new CultureInfo("ru-Ru"));
            WriteToFile(memoryStream1, filename);
	    }
		[Test]
		[Ignore("Слишком долго выполняется (но работает правильно)")]
		public void ExcelSimpleExportRowOverflow() {
		    const string filename = "TestSimpleOverflow.xls";
            DeleteTestFile(filename);
			var simpleTestData = NotRelectionTestDataEntities.CreateSimpleTestRowOverflowData();
			var firstTestDataRow = simpleTestData.Models.FirstOrDefault();
			Assert.NotNull(firstTestDataRow);
			var memoryStream = TableWriterSimple.Write(new XlsTableWriterSimple(), simpleTestData.Models, simpleTestData.ConfigurationBuilder.Value);
			WriteToFile(memoryStream, filename);
			var firstContact = firstTestDataRow.Contacts.FirstOrDefault();
			Assert.NotNull(firstContact);

			var firstContract = firstTestDataRow.Contracts.FirstOrDefault();
			Assert.NotNull(firstContract);

			var firstProduct = firstTestDataRow.Products.FirstOrDefault();
			Assert.NotNull(firstProduct);

			var firstEnumProp1 = firstTestDataRow.EnumProps1.FirstOrDefault();
			Assert.NotNull(firstEnumProp1);

			var firstEnumProp2 = firstTestDataRow.EnumProps2.FirstOrDefault();
			Assert.NotNull(firstEnumProp2);

			var parentColumnsQuantity = simpleTestData.ConfigurationBuilder.Value.ColumnsMap.Count;
			var contactsFieldsQuantity = firstContact.GetType().GetProperties().Count();
			var contractsFieldsQuantity = firstContract.GetType().GetProperties().Count();
			var productsFieldsQuantity = firstProduct.GetType().GetProperties().Count();
			var enumProp1FieldsQuantity = firstEnumProp1.GetType().GetProperties().Count();

			var testDataContactColumnsQuantity = simpleTestData.Models.Max(x => x.Contacts.Count) * contactsFieldsQuantity;
			var testDataContractColumnsQuantity = simpleTestData.Models.Max(x => x.Contracts.Count) * contractsFieldsQuantity;
			var testDataProductColumnsQuantity = simpleTestData.Models.Max(x => x.Products.Count) * productsFieldsQuantity;
			var testDataEnumProp1ColumnsQuantity = simpleTestData.Models.Max(x => x.EnumProps1.Count) * enumProp1FieldsQuantity;

			var parentColumnsNames = simpleTestData.ConfigurationBuilder.Value.ColumnsMap.Keys.ToList();
			var childsColumnsNames = simpleTestData.ConfigurationBuilder.Value.ChildrenMap.Select(childs => childs.ColumnsMap.Keys.ToList()).ToList();
			HSSFWorkbook xlsFile;
			using (var file = new FileStream(filename, FileMode.Open, FileAccess.Read)) {
				xlsFile = new HSSFWorkbook(file);
			}

			var header = parentColumnsNames.ToList();
			for (var childNumber = 0; childNumber < childsColumnsNames.Count; childNumber++) {
				var child = childsColumnsNames[childNumber];
				var childColumnsQuantity = 0;
				switch (childNumber) {
					case 0: {
							childColumnsQuantity = simpleTestData.Models.Max(x => x.Contacts.Count);
							break;
						}
					case 1: {
							childColumnsQuantity = simpleTestData.Models.Max(x => x.Contracts.Count);
							break;
						}
					case 2: {
							childColumnsQuantity = simpleTestData.Models.Max(x => x.Products.Count);
							break;
						}
					case 3: {
							childColumnsQuantity = simpleTestData.Models.Max(x => x.EnumProps1.Count);
							break;
						}
					case 4: {
							childColumnsQuantity = simpleTestData.Models.Max(x => x.EnumProps2.Count);
							break;
						}
				}
				for (var index = 1; index <= childColumnsQuantity; index++) {
					header.AddRange(child.Select(childPropertyName => childPropertyName + index));
				}
			}
			var modelNumber = 0;
			for (var sheetNumber = 0; sheetNumber < xlsFile.NumberOfSheets; sheetNumber++) {
				var sheet = xlsFile.GetSheetAt(sheetNumber);
				var row = sheet.GetRow(0);
				for (var columnNumber = 0; columnNumber < row.LastCellNum; columnNumber++) {
					Assert.AreEqual(row.GetCell(columnNumber).StringCellValue, header[columnNumber]);
				}
				for (var rowNumber = 1; rowNumber < sheet.LastRowNum; rowNumber++) {
					row = sheet.GetRow(rowNumber);
					Assert.NotNull(row);
					var currentTestDataRow = simpleTestData.Models[modelNumber];
					Assert.AreEqual(row.GetCell(0).StringCellValue, currentTestDataRow.Title ?? string.Empty);
					Assert.AreEqual(DateUtil.GetJavaDate(row.GetCell(1).NumericCellValue).ToString(CultureInfo.InvariantCulture), currentTestDataRow.RegistrationDate.ToString(CultureInfo.InvariantCulture));
					Assert.AreEqual(row.GetCell(2).StringCellValue, currentTestDataRow.Phone);
					Assert.AreEqual(row.GetCell(3).StringCellValue, currentTestDataRow.Inn);
					Assert.AreEqual(row.GetCell(4).StringCellValue, currentTestDataRow.Okato);

					Assert.AreEqual(row.GetCell(5).NumericCellValue, Convert.ToDouble(currentTestDataRow.Revenue));
					Assert.AreEqual(row.GetCell(6).NumericCellValue, Convert.ToDouble(currentTestDataRow.EmployeeCount));
					Assert.AreEqual(row.GetCell(7).StringCellValue, currentTestDataRow.IsActive.ToRussianString());
					Assert.AreEqual(row.GetCell(8).NumericCellValue, Convert.ToDouble(currentTestDataRow.Prop1));
					Assert.AreEqual(row.GetCell(9).StringCellValue, currentTestDataRow.Prop2);

					Assert.AreEqual(row.GetCell(10).StringCellValue, currentTestDataRow.Prop3.ToRussianString());
					Assert.AreEqual(row.GetCell(11).NumericCellValue, Convert.ToDouble(currentTestDataRow.Prop4));
					Assert.AreEqual(row.GetCell(12).StringCellValue, currentTestDataRow.Prop5);
					Assert.AreEqual(row.GetCell(13).StringCellValue, currentTestDataRow.Prop6.ToRussianString());
					Assert.AreEqual(row.GetCell(14).NumericCellValue, Convert.ToDouble(currentTestDataRow.Prop7));

					for (var contactNumber = 0; contactNumber < currentTestDataRow.Contacts.Count; contactNumber++) {
						var currentContactRow = currentTestDataRow.Contacts[contactNumber];
						Assert.AreEqual(row.GetCell(parentColumnsQuantity + (contactNumber * contactsFieldsQuantity)).StringCellValue, currentContactRow.Title);
						Assert.AreEqual(row.GetCell(parentColumnsQuantity + (contactNumber * contactsFieldsQuantity + 1)).StringCellValue, currentContactRow.Email);
					}

					for (var contractNumber = 0; contractNumber < currentTestDataRow.Contracts.Count; contractNumber++) {
						Assert.AreEqual(DateUtil.GetJavaDate(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + (contractNumber * contractsFieldsQuantity)).NumericCellValue).ToString(CultureInfo.InvariantCulture),
							currentTestDataRow.Contracts[contractNumber].BeginDate.ToString(CultureInfo.InvariantCulture));
						Assert.AreEqual(DateUtil.GetJavaDate(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + (contractNumber * contractsFieldsQuantity + 1)).NumericCellValue).ToString(CultureInfo.InvariantCulture),
							currentTestDataRow.Contracts[contractNumber].EndDate.ToString(CultureInfo.InvariantCulture));
						Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + (contractNumber * contractsFieldsQuantity + 2)).StringCellValue,
							currentTestDataRow.Contracts[contractNumber].Status.ToRussianString());
					}

					for (var productNumber = 0; productNumber < currentTestDataRow.Products.Count; productNumber++) {
						Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + (productNumber * 2)).StringCellValue,
							currentTestDataRow.Products[productNumber].Title);
						Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + (productNumber * 2 + 1)).NumericCellValue,
							Convert.ToDouble(currentTestDataRow.Products[productNumber].Amount));
					}

					for (var enumProp1Number = 0; enumProp1Number < currentTestDataRow.EnumProps1.Count; enumProp1Number++) {
						Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + (enumProp1Number * 3)).StringCellValue,
							currentTestDataRow.EnumProps1[enumProp1Number].Field1);
						Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + (enumProp1Number * 3 + 1)).NumericCellValue,
							Convert.ToDouble(currentTestDataRow.EnumProps1[enumProp1Number].Field2));
						Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + (enumProp1Number * 3 + 2)).StringCellValue,
							currentTestDataRow.EnumProps1[enumProp1Number].Field3.ToRussianString());
					}

					for (var enumProp2Number = 0; enumProp2Number < currentTestDataRow.EnumProps2.Count; enumProp2Number++) {
						Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + testDataEnumProp1ColumnsQuantity + (enumProp2Number * 5)).StringCellValue,
							currentTestDataRow.EnumProps2[enumProp2Number].Field4);
						Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + testDataEnumProp1ColumnsQuantity + (enumProp2Number * 5 + 1)).NumericCellValue,
							Convert.ToDouble(currentTestDataRow.EnumProps2[enumProp2Number].Field5));
						Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + testDataEnumProp1ColumnsQuantity + (enumProp2Number * 5 + 2)).StringCellValue,
							currentTestDataRow.EnumProps2[enumProp2Number].Field6.ToRussianString());
						Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + testDataEnumProp1ColumnsQuantity + (enumProp2Number * 5 + 3)).StringCellValue,
							currentTestDataRow.EnumProps2[enumProp2Number].Field7);
						Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + testDataEnumProp1ColumnsQuantity + (enumProp2Number * 5 + 4)).NumericCellValue,
							Convert.ToDouble(currentTestDataRow.EnumProps2[enumProp2Number].Field8));
					}
					modelNumber++;
				}
				modelNumber++;
			}
		}
		[Test]
		[Ignore("Слишком долго выполняется (но работает правильно)")]
		public void ExcelComplexExportRowOverflow() {
		    const string filename = "TestComplexOverflow.xls";
            DeleteTestFile(filename);
			var simpleTestData = NotRelectionTestDataEntities.CreateSimpleTestRowOverflowData();
			var firstTestDataRow = simpleTestData.Models.FirstOrDefault();
			Assert.NotNull(firstTestDataRow);
			var style = new TableWriterStyle();
			var memoryStream = TableWriterComplex.Write(new XlsTableWriterComplex(style), simpleTestData.Models, simpleTestData.ConfigurationBuilder.Value);
			WriteToFile(memoryStream, filename);

			HSSFWorkbook xlsFile;
			using (var file = new FileStream(filename, FileMode.Open, FileAccess.Read)) {
				xlsFile = new HSSFWorkbook(file);
			}
			var parentColumnsNames = simpleTestData.ConfigurationBuilder.Value.ColumnsMap.Keys.ToList();
			var sheetNumber = 0;
			var modelsQuantity = simpleTestData.Models.Count;
			var rowNumber = 0;

			var sheet = xlsFile.GetSheetAt(sheetNumber);
            var lastChildHeader = new List<string>();
			for (var modelNumber = 0; modelNumber < modelsQuantity; modelNumber++) {
				var lastParentHeader = new List<string>();
				
				var currentTestDataRow = simpleTestData.Models[modelNumber];
				var row = sheet.GetRow(rowNumber);
				var cell = row.GetCell(0);
				Assert.AreEqual(cell.StringCellValue, simpleTestData.ConfigurationBuilder.Value.Title);
				for (var cellNumber = 1; cellNumber < row.LastCellNum; cellNumber++) {
					cell = row.GetCell(cellNumber);
					Assert.AreEqual(cell.StringCellValue, parentColumnsNames[cellNumber - 1]);
					lastParentHeader.Add(parentColumnsNames[cellNumber - 1]);
				}
				RowNumberIncrement(lastParentHeader, lastChildHeader, xlsFile, ref rowNumber, ref sheet, ref sheetNumber);
				row = sheet.GetRow(rowNumber);

				Assert.AreEqual(row.GetCell(1).StringCellValue, currentTestDataRow.Title ?? string.Empty);
                Assert.AreEqual(DateUtil.GetJavaDate(row.GetCell(2).NumericCellValue).ToString(CultureInfo.InvariantCulture), currentTestDataRow.RegistrationDate.ToString(CultureInfo.InvariantCulture));
				Assert.AreEqual(row.GetCell(3).StringCellValue, currentTestDataRow.Phone);
				Assert.AreEqual(row.GetCell(4).StringCellValue, currentTestDataRow.Inn);
				Assert.AreEqual(row.GetCell(5).StringCellValue, currentTestDataRow.Okato);
				Assert.AreEqual(row.GetCell(6).NumericCellValue, Convert.ToDouble(currentTestDataRow.Revenue));
				Assert.AreEqual(row.GetCell(7).NumericCellValue, Convert.ToDouble(currentTestDataRow.EmployeeCount));
				Assert.AreEqual(row.GetCell(8).StringCellValue, currentTestDataRow.IsActive.ToRussianString());
				Assert.AreEqual(row.GetCell(9).NumericCellValue, Convert.ToDouble(currentTestDataRow.Prop1));
				Assert.AreEqual(row.GetCell(10).StringCellValue, currentTestDataRow.Prop2);
				Assert.AreEqual(row.GetCell(11).StringCellValue, currentTestDataRow.Prop3.ToRussianString());
				Assert.AreEqual(row.GetCell(12).NumericCellValue, Convert.ToDouble(currentTestDataRow.Prop4));
				Assert.AreEqual(row.GetCell(13).StringCellValue, currentTestDataRow.Prop5);
				Assert.AreEqual(row.GetCell(14).StringCellValue, currentTestDataRow.Prop6.ToRussianString());
				Assert.AreEqual(row.GetCell(15).NumericCellValue, Convert.ToDouble(currentTestDataRow.Prop7));
				RowNumberIncrement(lastParentHeader, null, xlsFile, ref rowNumber, ref sheet, ref sheetNumber);
				row = sheet.GetRow(rowNumber);

				var childsQuantity = simpleTestData.ConfigurationBuilder.Value.ChildrenMap.Count;

				for (var childNumber = 0; childNumber < childsQuantity; childNumber++) {
					switch (childNumber) {
						case 0: {
								var child = currentTestDataRow.Contacts;
								lastChildHeader = TestChildHeader(row, childNumber, simpleTestData);
                                RowNumberIncrement(lastParentHeader, lastChildHeader, xlsFile, ref rowNumber, ref sheet, ref sheetNumber);
								row = sheet.GetRow(rowNumber);
								foreach (var childProperty in child) {
									Assert.AreEqual(row.GetCell(1).StringCellValue, childProperty.Title);
									Assert.AreEqual(row.GetCell(2).StringCellValue, childProperty.Email);
									RowNumberIncrement(lastParentHeader, lastChildHeader, xlsFile, ref rowNumber, ref sheet, ref sheetNumber); 
									row = sheet.GetRow(rowNumber);
                                }
                                    lastChildHeader = null;
								break;
							}
						case 1: {
								var child = currentTestDataRow.Contracts;
								lastChildHeader = TestChildHeader(row, childNumber, simpleTestData);
								RowNumberIncrement(lastParentHeader, lastChildHeader, xlsFile, ref rowNumber, ref sheet, ref sheetNumber);
								row = sheet.GetRow(rowNumber);
								foreach (var childProperty in child) {
									Assert.AreEqual(DateUtil.GetJavaDate(row.GetCell(1).NumericCellValue).ToString(CultureInfo.InvariantCulture), childProperty.BeginDate.ToString(CultureInfo.InvariantCulture));
									Assert.AreEqual(DateUtil.GetJavaDate(row.GetCell(2).NumericCellValue).ToString(CultureInfo.InvariantCulture), childProperty.EndDate.ToString(CultureInfo.InvariantCulture));
									Assert.AreEqual(row.GetCell(3).StringCellValue, childProperty.Status.ToRussianString());
									RowNumberIncrement(lastParentHeader, lastChildHeader, xlsFile, ref rowNumber, ref sheet, ref sheetNumber); 
									row = sheet.GetRow(rowNumber);
								}
                                lastChildHeader = null;
								break;
							}
						case 2: {
								var child = currentTestDataRow.Products;
								lastChildHeader = TestChildHeader(row, childNumber, simpleTestData);
								RowNumberIncrement(lastParentHeader, lastChildHeader, xlsFile, ref rowNumber, ref sheet, ref sheetNumber);
								row = sheet.GetRow(rowNumber);
								foreach (var childProperty in child) {
									Assert.AreEqual(row.GetCell(1).StringCellValue, childProperty.Title);
									Assert.AreEqual(row.GetCell(2).NumericCellValue, Convert.ToDouble(childProperty.Amount));
									RowNumberIncrement(lastParentHeader, lastChildHeader, xlsFile, ref rowNumber, ref sheet, ref sheetNumber); 
									row = sheet.GetRow(rowNumber);
								}
                                lastChildHeader = null;
								break;
							}
						case 3: {
								var child = currentTestDataRow.EnumProps1;
								lastChildHeader = TestChildHeader(row, childNumber, simpleTestData);
								RowNumberIncrement(lastParentHeader, lastChildHeader, xlsFile, ref rowNumber, ref sheet, ref sheetNumber);
								row = sheet.GetRow(rowNumber);
								foreach (var childProperty in child) {
									Assert.AreEqual(row.GetCell(1).StringCellValue, childProperty.Field1);
									Assert.AreEqual(row.GetCell(2).NumericCellValue, Convert.ToDouble(childProperty.Field2));
									Assert.AreEqual(row.GetCell(3).StringCellValue, childProperty.Field3.ToRussianString());
									RowNumberIncrement(lastParentHeader, lastChildHeader, xlsFile, ref rowNumber, ref sheet, ref sheetNumber); 
									row = sheet.GetRow(rowNumber);
								}
                                lastChildHeader = null;
								break;
							}
						case 4: {
								var child = currentTestDataRow.EnumProps2;
								lastChildHeader = TestChildHeader(row, childNumber, simpleTestData);
								RowNumberIncrement(lastParentHeader, lastChildHeader, xlsFile, ref rowNumber, ref sheet, ref sheetNumber);
								row = sheet.GetRow(rowNumber);
								foreach (var childProperty in child) {
									Assert.AreEqual(row.GetCell(1).StringCellValue, childProperty.Field4);
									Assert.AreEqual(row.GetCell(2).NumericCellValue, Convert.ToDouble(childProperty.Field5));
									Assert.AreEqual(row.GetCell(3).StringCellValue, childProperty.Field6.ToRussianString());
									Assert.AreEqual(row.GetCell(4).StringCellValue, childProperty.Field7);
									Assert.AreEqual(row.GetCell(5).NumericCellValue, Convert.ToDouble(childProperty.Field8));
									RowNumberIncrement(lastParentHeader, lastChildHeader, xlsFile, ref rowNumber, ref sheet, ref sheetNumber); 
									row = sheet.GetRow(rowNumber);
								}
                                lastChildHeader = null;
								break;
							}
					}
				}
			}
		}
		private static void RowNumberIncrement(IReadOnlyList<string> lastParentHeader, List<string> lastChildHeader, IWorkbook workbook, ref int rowNumber, ref ISheet sheet, ref int sheetNumber, bool lastCellIsChildHeader = false) {
			if (rowNumber >= 65535) {
				sheetNumber++;
				sheet = workbook.GetSheetAt(sheetNumber);
				var row = sheet.GetRow(0);
				for (var columnNumber = 1; columnNumber < row.LastCellNum; columnNumber++) {
					Assert.AreEqual(row.GetCell(columnNumber).StringCellValue, lastParentHeader[columnNumber - 1]);
				}
				row = sheet.GetRow(1);
                if (lastChildHeader != null) {
					for (var columnNumber = 1; columnNumber < row.LastCellNum; columnNumber++) {
						Assert.AreEqual(row.GetCell(columnNumber).StringCellValue, lastChildHeader[columnNumber - 1]);
					}
					rowNumber = 2;
				} else {
					rowNumber = 1;
				}
			} else {
				rowNumber++;
			}
		}
		[Test]
		public void ExcelSimpleExport() {
		    const string filename = "TestSimple.xls";
            DeleteTestFile(filename);
			var simpleTestData = NotRelectionTestDataEntities.CreateSimpleTestData();
			var firstTestDataRow = simpleTestData.Models.FirstOrDefault();
            Assert.NotNull(firstTestDataRow);
            var memoryStream = TableWriterSimple.Write(new XlsTableWriterSimple(), simpleTestData.Models, simpleTestData.ConfigurationBuilder.Value);
            WriteToFile(memoryStream, filename);

            var firstContact = firstTestDataRow.Contacts.FirstOrDefault();
            Assert.NotNull(firstContact);

            var firstContract = firstTestDataRow.Contracts.FirstOrDefault();
            Assert.NotNull(firstContract);

            var firstProduct = firstTestDataRow.Products.FirstOrDefault();
            Assert.NotNull(firstProduct);

			var firstEnumProp1 = firstTestDataRow.EnumProps1.FirstOrDefault();
            Assert.NotNull(firstEnumProp1);

			var firstEnumProp2 = firstTestDataRow.EnumProps2.FirstOrDefault();
            Assert.NotNull(firstEnumProp2);

			var parentColumnsQuantity = simpleTestData.ConfigurationBuilder.Value.ColumnsMap.Count;
			var contactsFieldsQuantity = firstContact.GetType().GetProperties().Count();
			var contractsFieldsQuantity = firstContract.GetType().GetProperties().Count();
			var productsFieldsQuantity = firstProduct.GetType().GetProperties().Count();
			var enumProp1FieldsQuantity = firstEnumProp1.GetType().GetProperties().Count();

			var testDataContactColumnsQuantity = simpleTestData.Models.Max(x => x.Contacts.Count) * contactsFieldsQuantity;
			var testDataContractColumnsQuantity = simpleTestData.Models.Max(x => x.Contracts.Count) * contractsFieldsQuantity;
			var testDataProductColumnsQuantity = simpleTestData.Models.Max(x => x.Products.Count) * productsFieldsQuantity;
			var testDataEnumProp1ColumnsQuantity = simpleTestData.Models.Max(x => x.EnumProps1.Count) * enumProp1FieldsQuantity;

			var parentColumnsNames = simpleTestData.ConfigurationBuilder.Value.ColumnsMap.Keys.ToList();
			var childsColumnsNames = simpleTestData.ConfigurationBuilder.Value.ChildrenMap.Select(childs => childs.ColumnsMap.Keys.ToList()).ToList();
			HSSFWorkbook xlsFile;
			using (var file = new FileStream(filename, FileMode.Open, FileAccess.Read)) {
				xlsFile = new HSSFWorkbook(file);
			}
			var sheet = xlsFile.GetSheetAt(0);
			var rowNumber = 0;
			var row = sheet.GetRow(rowNumber);

			int columnNumber;
			for (columnNumber = 0; columnNumber < parentColumnsNames.Count; columnNumber++) {
				Assert.AreEqual(row.GetCell(columnNumber).StringCellValue, parentColumnsNames[columnNumber]);
			}

			columnNumber = parentColumnsQuantity;
			for (var childNumber = 0; childNumber < childsColumnsNames.Count; childNumber++) {
				var child = childsColumnsNames[childNumber];
				var childColumnsQuantity = 0;
				switch (childNumber) {
					case 0: {
							childColumnsQuantity = simpleTestData.Models.Max(x => x.Contacts.Count);
							break;
						}
					case 1: {
							childColumnsQuantity = simpleTestData.Models.Max(x => x.Contracts.Count);
							break;
						}
					case 2: {
							childColumnsQuantity = simpleTestData.Models.Max(x => x.Products.Count);
							break;
						}
					case 3: {
							childColumnsQuantity = simpleTestData.Models.Max(x => x.EnumProps1.Count);
							break;
						}
					case 4: {
							childColumnsQuantity = simpleTestData.Models.Max(x => x.EnumProps2.Count);
							break;
						}
				}
				for (var index = 1; index <= childColumnsQuantity; index++) {
					foreach (var childPropertyName in child) {
						Assert.AreEqual(row.GetCell(columnNumber).StringCellValue, childPropertyName + index);
						columnNumber++;
					}
				}
			}

			for (rowNumber = 1; rowNumber <= sheet.LastRowNum; rowNumber++) {
				row = sheet.GetRow(rowNumber);
				Assert.NotNull(row);
				var currentTestDataRow = simpleTestData.Models[rowNumber - 1];
				Assert.AreEqual(row.GetCell(0).StringCellValue, currentTestDataRow.Title);
                Assert.AreEqual(DateUtil.GetJavaDate(row.GetCell(1).NumericCellValue).ToString(CultureInfo.InvariantCulture), currentTestDataRow.RegistrationDate.ToString(CultureInfo.InvariantCulture));
				Assert.AreEqual(row.GetCell(2).StringCellValue, currentTestDataRow.Phone);
				Assert.AreEqual(row.GetCell(3).StringCellValue, currentTestDataRow.Inn);
				Assert.AreEqual(row.GetCell(4).StringCellValue, currentTestDataRow.Okato);

				Assert.AreEqual(row.GetCell(5).NumericCellValue, Convert.ToDouble(currentTestDataRow.Revenue));
				Assert.AreEqual(row.GetCell(6).NumericCellValue, Convert.ToDouble(currentTestDataRow.EmployeeCount));
				Assert.AreEqual(row.GetCell(7).StringCellValue, currentTestDataRow.IsActive.ToRussianString());
				Assert.AreEqual(row.GetCell(8).NumericCellValue, Convert.ToDouble(currentTestDataRow.Prop1));
				Assert.AreEqual(row.GetCell(9).StringCellValue, currentTestDataRow.Prop2);

				Assert.AreEqual(row.GetCell(10).StringCellValue, currentTestDataRow.Prop3.ToRussianString());
				Assert.AreEqual(row.GetCell(11).NumericCellValue, Convert.ToDouble(currentTestDataRow.Prop4));
				Assert.AreEqual(row.GetCell(12).StringCellValue, currentTestDataRow.Prop5);
				Assert.AreEqual(row.GetCell(13).StringCellValue, currentTestDataRow.Prop6.ToRussianString());
				Assert.AreEqual(row.GetCell(14).NumericCellValue, Convert.ToDouble(currentTestDataRow.Prop7));

				for (var contactNumber = 0; contactNumber < currentTestDataRow.Contacts.Count; contactNumber++) {
					var currentContactRow = currentTestDataRow.Contacts[contactNumber];
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + (contactNumber * contactsFieldsQuantity)).StringCellValue, currentContactRow.Title);
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + (contactNumber * contactsFieldsQuantity + 1)).StringCellValue, currentContactRow.Email);
				}

				for (var contractNumber = 0; contractNumber < currentTestDataRow.Contracts.Count; contractNumber++) {
					Assert.AreEqual(DateUtil.GetJavaDate(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + (contractNumber * contractsFieldsQuantity)).NumericCellValue).ToString(CultureInfo.InvariantCulture),
                        currentTestDataRow.Contracts[contractNumber].BeginDate.ToString(CultureInfo.InvariantCulture));
					Assert.AreEqual(DateUtil.GetJavaDate(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + (contractNumber * contractsFieldsQuantity + 1)).NumericCellValue).ToString(CultureInfo.InvariantCulture),
                        currentTestDataRow.Contracts[contractNumber].EndDate.ToString(CultureInfo.InvariantCulture));
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + (contractNumber * contractsFieldsQuantity + 2)).StringCellValue,
						currentTestDataRow.Contracts[contractNumber].Status.ToRussianString());
				}

				for (var productNumber = 0; productNumber < currentTestDataRow.Products.Count; productNumber++) {
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + (productNumber * 2)).StringCellValue,
						currentTestDataRow.Products[productNumber].Title);
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + (productNumber * 2 + 1)).NumericCellValue,
						Convert.ToDouble(currentTestDataRow.Products[productNumber].Amount));
				}

				for (var enumProp1Number = 0; enumProp1Number < currentTestDataRow.EnumProps1.Count; enumProp1Number++) {
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + (enumProp1Number * 3)).StringCellValue,
						currentTestDataRow.EnumProps1[enumProp1Number].Field1);
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + (enumProp1Number * 3 + 1)).NumericCellValue,
						Convert.ToDouble(currentTestDataRow.EnumProps1[enumProp1Number].Field2));
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + (enumProp1Number * 3 + 2)).StringCellValue,
						currentTestDataRow.EnumProps1[enumProp1Number].Field3.ToRussianString());
				}

				for (var enumProp2Number = 0; enumProp2Number < currentTestDataRow.EnumProps2.Count; enumProp2Number++) {
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + testDataEnumProp1ColumnsQuantity + (enumProp2Number * 5)).StringCellValue,
						currentTestDataRow.EnumProps2[enumProp2Number].Field4);
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + testDataEnumProp1ColumnsQuantity + (enumProp2Number * 5 + 1)).NumericCellValue,
						Convert.ToDouble(currentTestDataRow.EnumProps2[enumProp2Number].Field5));
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + testDataEnumProp1ColumnsQuantity + (enumProp2Number * 5 + 2)).StringCellValue,
						currentTestDataRow.EnumProps2[enumProp2Number].Field6.ToRussianString());
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + testDataEnumProp1ColumnsQuantity + (enumProp2Number * 5 + 3)).StringCellValue,
						currentTestDataRow.EnumProps2[enumProp2Number].Field7);
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + testDataEnumProp1ColumnsQuantity + (enumProp2Number * 5 + 4)).NumericCellValue,
						Convert.ToDouble(currentTestDataRow.EnumProps2[enumProp2Number].Field8));
				}
			}
		}
		[Test]
		public void ExcelStyleSimpleExport() {
		    const string filename = "TestStyleSimple.xls";
            DeleteTestFile(filename);
			var simpleTestData = NotRelectionTestDataEntities.CreateSimpleTestData(true);
			var firstTestDataRow = simpleTestData.Models.FirstOrDefault();
			Assert.NotNull(firstTestDataRow);
			var style = new TableWriterStyle();
			var memoryStream = TableWriterSimple.Write(new XlsTableWriterSimple(style), simpleTestData.Models, simpleTestData.ConfigurationBuilder.Value);
			WriteToFile(memoryStream, filename);
			HSSFWorkbook xlsFile;
			using (var file = new FileStream(filename, FileMode.Open, FileAccess.Read)) {
				xlsFile = new HSSFWorkbook(file);
			}
			var sheet = xlsFile.GetSheetAt(0);

			short[] red = { 255, 0, 0 };
			short[] green = { 0, 255, 0 };
			short[] blue = { 0, 0, 255 };

			var rowNumber = 0;
			var row = sheet.GetRow(rowNumber);
			for (var cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
				CustomAssert.IsEqualFont(xlsFile, sheet, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Bold);
			}

			for (rowNumber = 1; rowNumber < sheet.LastRowNum; rowNumber++) {
				row = sheet.GetRow(rowNumber);
				for (var cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
					CustomAssert.IsEqualFont(xlsFile, sheet, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Normal);
				}
			}
			CustomAssert.IsEqualExcelColor((HSSFColor)sheet.GetRow(1).GetCell(2).CellStyle.FillForegroundColorColor, green);
			CustomAssert.IsEqualExcelColor((HSSFColor)sheet.GetRow(2).GetCell(0).CellStyle.FillForegroundColorColor, red);
		}
		[Test]
		public void ExcelComplexExport() {
		    const string filename = "TestComplex.xls";
            DeleteTestFile(filename);
			var simpleTestData = NotRelectionTestDataEntities.CreateSimpleTestData();
			var firstTestDataRow = simpleTestData.Models.FirstOrDefault();
			Assert.NotNull(firstTestDataRow);
			var style = new TableWriterStyle();
			var memoryStream = TableWriterComplex.Write(new XlsTableWriterComplex(style), simpleTestData.Models, simpleTestData.ConfigurationBuilder.Value);
			WriteToFile(memoryStream, filename);

			HSSFWorkbook xlsFile;
			using (var file = new FileStream(filename, FileMode.Open, FileAccess.Read)) {
				xlsFile = new HSSFWorkbook(file);
			}
			var sheet = xlsFile.GetSheetAt(0);
			var parentColumnsNames = simpleTestData.ConfigurationBuilder.Value.ColumnsMap.Keys.ToList();

			var modelsQuantity = simpleTestData.Models.Count;
			var rowNumber = 0;
			for (var modelNumber = 0; modelNumber < modelsQuantity; modelNumber++) {
				var currentTestDataRow = simpleTestData.Models[modelNumber];
				var row = sheet.GetRow(rowNumber);
				var cell = row.GetCell(0);
				Assert.AreEqual(cell.StringCellValue, simpleTestData.ConfigurationBuilder.Value.Title);
				for (var cellNumber = 1; cellNumber < row.LastCellNum; cellNumber++) {
					cell = row.GetCell(cellNumber);
					Assert.AreEqual(cell.StringCellValue, parentColumnsNames[cellNumber - 1]);
				}
				rowNumber++;
				row = sheet.GetRow(rowNumber);

				Assert.AreEqual(row.GetCell(1).StringCellValue, currentTestDataRow.Title);
				Assert.AreEqual(DateUtil.GetJavaDate(row.GetCell(2).NumericCellValue).ToString(CultureInfo.InvariantCulture),  currentTestDataRow.RegistrationDate.ToString(CultureInfo.InvariantCulture));
				Assert.AreEqual(row.GetCell(3).StringCellValue, currentTestDataRow.Phone);
				Assert.AreEqual(row.GetCell(4).StringCellValue, currentTestDataRow.Inn);
				Assert.AreEqual(row.GetCell(5).StringCellValue, currentTestDataRow.Okato);
				Assert.AreEqual(row.GetCell(6).NumericCellValue, Convert.ToDouble(currentTestDataRow.Revenue));
				Assert.AreEqual(row.GetCell(7).NumericCellValue, Convert.ToDouble(currentTestDataRow.EmployeeCount));
				Assert.AreEqual(row.GetCell(8).StringCellValue, currentTestDataRow.IsActive.ToRussianString());
				Assert.AreEqual(row.GetCell(9).NumericCellValue, Convert.ToDouble(currentTestDataRow.Prop1));
				Assert.AreEqual(row.GetCell(10).StringCellValue, currentTestDataRow.Prop2);

				Assert.AreEqual(row.GetCell(11).StringCellValue, currentTestDataRow.Prop3.ToRussianString());
				Assert.AreEqual(row.GetCell(12).NumericCellValue, Convert.ToDouble(currentTestDataRow.Prop4));
				Assert.AreEqual(row.GetCell(13).StringCellValue, currentTestDataRow.Prop5);
				Assert.AreEqual(row.GetCell(14).StringCellValue, currentTestDataRow.Prop6.ToRussianString());
				Assert.AreEqual(row.GetCell(15).NumericCellValue, Convert.ToDouble(currentTestDataRow.Prop7));
				rowNumber++;
				row = sheet.GetRow(rowNumber);

				var childsQuantity = simpleTestData.ConfigurationBuilder.Value.ChildrenMap.Count;

				for (var childNumber = 0; childNumber < childsQuantity; childNumber++) {
					switch (childNumber) {
						case 0: {
								var child = currentTestDataRow.Contacts;
								TestChildHeader(row, childNumber, simpleTestData);
								rowNumber++;
								row = sheet.GetRow(rowNumber);
								foreach (var childProperty in child) {
									Assert.AreEqual(row.GetCell(1).StringCellValue, childProperty.Title);
									Assert.AreEqual(row.GetCell(2).StringCellValue, childProperty.Email);
									rowNumber++;
									row = sheet.GetRow(rowNumber);
								}
								break;
							}
						case 1: {
								var child = currentTestDataRow.Contracts;
								TestChildHeader(row, childNumber, simpleTestData);
								rowNumber++;
								row = sheet.GetRow(rowNumber);
								foreach (var childProperty in child) {
                                    Assert.AreEqual(DateUtil.GetJavaDate(row.GetCell(1).NumericCellValue).ToString(CultureInfo.InvariantCulture), childProperty.BeginDate.ToString(CultureInfo.InvariantCulture));
                                    Assert.AreEqual(DateUtil.GetJavaDate(row.GetCell(2).NumericCellValue).ToString(CultureInfo.InvariantCulture), childProperty.EndDate.ToString(CultureInfo.InvariantCulture));
									Assert.AreEqual(row.GetCell(3).StringCellValue, childProperty.Status.ToRussianString());
									rowNumber++;
									row = sheet.GetRow(rowNumber);
								}
								break;
							}
						case 2: {
								var child = currentTestDataRow.Products;
								TestChildHeader(row, childNumber, simpleTestData);
								rowNumber++;
								row = sheet.GetRow(rowNumber);
								foreach (var childProperty in child) {
									Assert.AreEqual(row.GetCell(1).StringCellValue, childProperty.Title);
									Assert.AreEqual(row.GetCell(2).NumericCellValue, Convert.ToDouble(childProperty.Amount));
									rowNumber++;
									row = sheet.GetRow(rowNumber);
								}
								break;
							}
						case 3: {
								var child = currentTestDataRow.EnumProps1;
								TestChildHeader(row, childNumber, simpleTestData);
								rowNumber++;
								row = sheet.GetRow(rowNumber);
								foreach (var childProperty in child) {
									Assert.AreEqual(row.GetCell(1).StringCellValue, childProperty.Field1);
									Assert.AreEqual(row.GetCell(2).NumericCellValue, Convert.ToDouble(childProperty.Field2));
									Assert.AreEqual(row.GetCell(3).StringCellValue, childProperty.Field3.ToRussianString());
									rowNumber++;
									row = sheet.GetRow(rowNumber);
								}
								break;
							}
						case 4: {
								var child = currentTestDataRow.EnumProps2;
								TestChildHeader(row, childNumber, simpleTestData);
								rowNumber++;
								row = sheet.GetRow(rowNumber);
								foreach (var childProperty in child) {
									Assert.AreEqual(row.GetCell(1).StringCellValue, childProperty.Field4);
									Assert.AreEqual(row.GetCell(2).NumericCellValue, Convert.ToDouble(childProperty.Field5));
									Assert.AreEqual(row.GetCell(3).StringCellValue, childProperty.Field6.ToRussianString());
									Assert.AreEqual(row.GetCell(4).StringCellValue, childProperty.Field7);
									Assert.AreEqual(row.GetCell(5).NumericCellValue, Convert.ToDouble(childProperty.Field8));
									rowNumber++;
									row = sheet.GetRow(rowNumber);
								}
								break;
							}
					}
				}
			}
		}
		[Test]
		public void ExcelStyleComplexExport() {
		    const string filename = "TestStyleComplex.xls";
            DeleteTestFile(filename);
			var simpleTestData = NotRelectionTestDataEntities.CreateSimpleTestData(true);
			var firstTestDataRow = simpleTestData.Models.FirstOrDefault();
			Assert.NotNull(firstTestDataRow);
			var style = new TableWriterStyle();
			var memoryStream = TableWriterComplex.Write(new XlsTableWriterComplex(style), simpleTestData.Models, simpleTestData.ConfigurationBuilder.Value);
			WriteToFile(memoryStream, filename);

			HSSFWorkbook xlsFile;
			using (var file = new FileStream(filename, FileMode.Open, FileAccess.Read)) {
				xlsFile = new HSSFWorkbook(file);
			}
			var sheet = xlsFile.GetSheetAt(0);

			short[] red = { 255, 0, 0 };
			short[] green = { 0, 255, 0 };
			short[] blue = { 0, 0, 255 };

			CustomAssert.IsEqualExcelColor((HSSFColor)sheet.GetRow(1).GetCell(3).CellStyle.FillForegroundColorColor, green);
			CustomAssert.IsEqualExcelColor((HSSFColor)sheet.GetRow(3).GetCell(1).CellStyle.FillForegroundColorColor, blue);
			CustomAssert.IsEqualExcelColor((HSSFColor)sheet.GetRow(22).GetCell(1).CellStyle.FillForegroundColorColor, red);
			CustomAssert.IsEqualExcelColor((HSSFColor)sheet.GetRow(24).GetCell(1).CellStyle.FillForegroundColorColor, blue);

			var modelsQuantity = simpleTestData.Models.Count;
			var rowNumber = 0;
			var childsQuantity = 0;
			for (var modelNumber = 0; modelNumber < modelsQuantity; modelNumber++) {
				childsQuantity += 7 + simpleTestData.Models[modelNumber].Contacts.Count + simpleTestData.Models[modelNumber].Contracts.Count + simpleTestData.Models[modelNumber].Products.Count + simpleTestData.Models[modelNumber].EnumProps1.Count + simpleTestData.Models[modelNumber].EnumProps2.Count;
				var row = sheet.GetRow(rowNumber);
				for (var cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
					CustomAssert.IsEqualFont(xlsFile, sheet, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Bold);
				}
				rowNumber++;

				for (; rowNumber < childsQuantity; rowNumber++) {
					row = sheet.GetRow(rowNumber);
					for (var cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
						CustomAssert.IsEqualFont(xlsFile, sheet, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Normal);
					}
				}
			}
		}
		[Test]
		public void ExcelReflectionSimpleExport() {
            var filename = "TestReflectionSimple.xls";
            DeleteTestFile(filename);
		    var testData = TestDataEntities.CreateSimpleTestDataModels();
            var memoryStream1 = ReflectionWriterSimple.Write(testData, new XlsTableWriterSimple(), new CultureInfo("ru-Ru"));

            WriteToFile(memoryStream1, filename);
            ExcelReflectionSimpleExportTest(testData, filename);

            filename = "RandomTestReflectionSimple.xls";
            DeleteTestFile(filename);
			var test = new List<ReflectionTestDataEntities>();
			var rand = new Random();
			for (var i = 2; i < rand.Next(10, 30); i++) {
				test.Add(new ReflectionTestDataEntities());
			}
			var memoryStream = ReflectionWriterSimple.Write(test, new XlsTableWriterSimple(), new CultureInfo("ru-Ru"));
            
			WriteToFile(memoryStream, filename);
			ExcelReflectionSimpleExportTest(test, filename);
		}
		[Test]
		public void ExcelStyleReflectionSimpleExport() {
			var models = TestDataEntities.CreateSimpleTestDataModels();
			var style = new TableWriterStyle();
			var memoryStream = ReflectionWriterSimple.Write(models, new XlsTableWriterSimple(style), new CultureInfo("ru-Ru"));
			var filename = "TestReflectionStyleSimple.xls";
            DeleteTestFile(filename);
			WriteToFile(memoryStream, filename);

			ExcelStyleReflectionSimpleExportTest(models, filename);

			var test = new List<ReflectionTestDataEntities>();
			var rand = new Random();
			for (var i = 2; i < rand.Next(10, 30); i++) {
				test.Add(new ReflectionTestDataEntities());
			}
			memoryStream = ReflectionWriterSimple.Write(test, new XlsTableWriterSimple(), new CultureInfo("ru-Ru"));
		    filename = "RandomStyleTestReflectionSimple.xls";
            DeleteTestFile(filename);
			WriteToFile(memoryStream, filename);
			ExcelStyleReflectionSimpleExportTest(test, filename);
		}
		[Test]
		public void ExcelReflectionComplexExport() {
			var style = new TableWriterStyle();
			var test = new List<ReflectionTestDataEntities>();
			var rand = new Random();
			for (var i = 2; i < rand.Next(10, 30); i++) {
				test.Add(new ReflectionTestDataEntities());
			}
			var memoryStream = ReflectionWriterComplex.Write(test, new XlsTableWriterComplex(style), new CultureInfo("ru-Ru"));
		    var filename = "RandomTestReflectionComplex.xls";
            DeleteTestFile(filename);
			WriteToFile(memoryStream, filename);
			ExcelReflectionComplexExportTest(test, filename);
		}
		[Test]
		public void ExcelStyleReflectionComplexExport() {
			var models = TestDataEntities.CreateSimpleTestDataModels();
			var style = new TableWriterStyle();
			var memoryStream = ReflectionWriterComplex.Write(models, new XlsTableWriterComplex(style), new CultureInfo("ru-Ru"));
			var filename = "TestReflectionStyleComplex.xls";
            DeleteTestFile(filename);
			WriteToFile(memoryStream, filename);

			ExcelStyleReflectionComplexExportTest(models, filename);

			var test = new List<ReflectionTestDataEntities>();
			var rand = new Random();
			for (var i = 2; i < rand.Next(10, 30); i++) {
				test.Add(new ReflectionTestDataEntities());
			}
			memoryStream = ReflectionWriterComplex.Write(test, new XlsTableWriterComplex(style), new CultureInfo("ru-Ru"));
		    filename = "RandomStyleTestReflectionComplex.xls";
            DeleteTestFile(filename);
			WriteToFile(memoryStream, filename);
			ExcelStyleReflectionComplexExportTest(test, filename);
		}

		[Test]
		public void ExcelSimpleXlsxExport() {
		    const string filename = "TestSimple.xlsx";
            DeleteTestFile(filename);
			var simpleTestData = NotRelectionTestDataEntities.CreateSimpleTestData(true);
			var firstTestDataRow = simpleTestData.Models.FirstOrDefault();
			Assert.NotNull(firstTestDataRow);
			var style = new TableWriterStyle();
			var ms = TableWriterSimple.Write(new XlsxTableWriterSimple(style), simpleTestData.Models, simpleTestData.ConfigurationBuilder.Value);
            WriteToFile(ms, filename);
            var firstContact = firstTestDataRow.Contacts.FirstOrDefault();
			Assert.NotNull(firstContact);

			var firstContract = firstTestDataRow.Contracts.FirstOrDefault();
			Assert.NotNull(firstContract);

			var firstProduct = firstTestDataRow.Products.FirstOrDefault();
			Assert.NotNull(firstProduct);

			var firstEnumProp1 = firstTestDataRow.EnumProps1.FirstOrDefault();
			Assert.NotNull(firstEnumProp1);

			var firstEnumProp2 = firstTestDataRow.EnumProps2.FirstOrDefault();
			Assert.NotNull(firstEnumProp2);

			var parentColumnsQuantity = simpleTestData.ConfigurationBuilder.Value.ColumnsMap.Count;
			var contactsFieldsQuantity = firstContact.GetType().GetProperties().Count();
			var contractsFieldsQuantity = firstContract.GetType().GetProperties().Count();
			var productsFieldsQuantity = firstProduct.GetType().GetProperties().Count();
			var enumProp1FieldsQuantity = firstEnumProp1.GetType().GetProperties().Count();

			var testDataContactColumnsQuantity = simpleTestData.Models.Max(x => x.Contacts.Count) * contactsFieldsQuantity;
			var testDataContractColumnsQuantity = simpleTestData.Models.Max(x => x.Contracts.Count) * contractsFieldsQuantity;
			var testDataProductColumnsQuantity = simpleTestData.Models.Max(x => x.Products.Count) * productsFieldsQuantity;
			var testDataEnumProp1ColumnsQuantity = simpleTestData.Models.Max(x => x.EnumProps1.Count) * enumProp1FieldsQuantity;

			var parentColumnsNames = simpleTestData.ConfigurationBuilder.Value.ColumnsMap.Keys.ToList();
			var childsColumnsNames = simpleTestData.ConfigurationBuilder.Value.ChildrenMap.Select(childs => childs.ColumnsMap.Keys.ToList()).ToList();
			XSSFWorkbook xlsFile;
			using (var file = new FileStream(filename, FileMode.Open, FileAccess.Read)) {
				xlsFile = new XSSFWorkbook(file);
			}
			var sheet = xlsFile.GetSheetAt(0);
			var rowNumber = 0;
			var row = sheet.GetRow(rowNumber);

			int columnNumber;
			for (columnNumber = 0; columnNumber < parentColumnsNames.Count; columnNumber++) {
				Assert.AreEqual(row.GetCell(columnNumber).StringCellValue, parentColumnsNames[columnNumber]);
			}

			columnNumber = parentColumnsQuantity;
			for (var childNumber = 0; childNumber < childsColumnsNames.Count; childNumber++) {
				var child = childsColumnsNames[childNumber];
				var childColumnsQuantity = 0;
				switch (childNumber) {
					case 0: {
							childColumnsQuantity = simpleTestData.Models.Max(x => x.Contacts.Count);
							break;
						}
					case 1: {
							childColumnsQuantity = simpleTestData.Models.Max(x => x.Contracts.Count);
							break;
						}
					case 2: {
							childColumnsQuantity = simpleTestData.Models.Max(x => x.Products.Count);
							break;
						}
					case 3: {
							childColumnsQuantity = simpleTestData.Models.Max(x => x.EnumProps1.Count);
							break;
						}
					case 4: {
							childColumnsQuantity = simpleTestData.Models.Max(x => x.EnumProps2.Count);
							break;
						}
				}
				for (var index = 1; index <= childColumnsQuantity; index++) {
					foreach (var childPropertyName in child) {
						Assert.AreEqual(row.GetCell(columnNumber).StringCellValue, childPropertyName + index);
						columnNumber++;
					}
				}
			}
            for (rowNumber = 1; rowNumber <= sheet.LastRowNum; rowNumber++) {
                row = sheet.GetRow(rowNumber);
                Assert.NotNull(row);
                var currentTestDataRow = simpleTestData.Models[rowNumber - 1];
                Assert.AreEqual(row.GetCell(0).StringCellValue, currentTestDataRow.Title);
                Assert.AreEqual(row.GetCell(1).CellFormula, string.Format("=Date({0},{1},{2})", currentTestDataRow.RegistrationDate.Year, currentTestDataRow.RegistrationDate.Month, currentTestDataRow.RegistrationDate.Day));
                Assert.AreEqual(row.GetCell(2).StringCellValue, currentTestDataRow.Phone);
                Assert.AreEqual(row.GetCell(3).StringCellValue, currentTestDataRow.Inn);
                Assert.AreEqual(row.GetCell(4).StringCellValue, currentTestDataRow.Okato);

                Assert.AreEqual(row.GetCell(5).NumericCellValue, Convert.ToDouble(currentTestDataRow.Revenue));
                Assert.AreEqual(row.GetCell(6).NumericCellValue, Convert.ToDouble(currentTestDataRow.EmployeeCount));
                Assert.AreEqual(row.GetCell(7).StringCellValue, currentTestDataRow.IsActive.ToRussianString());
                Assert.AreEqual(row.GetCell(8).NumericCellValue, Convert.ToDouble(currentTestDataRow.Prop1));
                Assert.AreEqual(row.GetCell(9).StringCellValue, currentTestDataRow.Prop2);

                Assert.AreEqual(row.GetCell(10).StringCellValue, currentTestDataRow.Prop3.ToRussianString());
                Assert.AreEqual(row.GetCell(11).NumericCellValue, Convert.ToDouble(currentTestDataRow.Prop4));
                Assert.AreEqual(row.GetCell(12).StringCellValue, currentTestDataRow.Prop5);
                Assert.AreEqual(row.GetCell(13).StringCellValue, currentTestDataRow.Prop6.ToRussianString());
                Assert.AreEqual(row.GetCell(14).NumericCellValue, Convert.ToDouble(currentTestDataRow.Prop7));

                for (var contactNumber = 0; contactNumber < currentTestDataRow.Contacts.Count; contactNumber++) {
                    var currentContactRow = currentTestDataRow.Contacts[contactNumber];
                    Assert.AreEqual(row.GetCell(parentColumnsQuantity + (contactNumber * contactsFieldsQuantity)).StringCellValue, currentContactRow.Title);
                    Assert.AreEqual(row.GetCell(parentColumnsQuantity + (contactNumber * contactsFieldsQuantity + 1)).StringCellValue, currentContactRow.Email);
                }

                for (var contractNumber = 0; contractNumber < currentTestDataRow.Contracts.Count; contractNumber++) {
                    Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + (contractNumber * contractsFieldsQuantity)).CellFormula,
                        string.Format("=Date({0},{1},{2})", currentTestDataRow.Contracts[contractNumber].BeginDate.Year, currentTestDataRow.Contracts[contractNumber].BeginDate.Month, currentTestDataRow.Contracts[contractNumber].BeginDate.Day));
                    Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + (contractNumber * contractsFieldsQuantity + 1)).CellFormula,
                        string.Format("=Date({0},{1},{2})", currentTestDataRow.Contracts[contractNumber].EndDate.Year, currentTestDataRow.Contracts[contractNumber].EndDate.Month, currentTestDataRow.Contracts[contractNumber].EndDate.Day));
                    Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + (contractNumber * contractsFieldsQuantity + 2)).StringCellValue,
                        currentTestDataRow.Contracts[contractNumber].Status.ToRussianString());
                }

                for (var productNumber = 0; productNumber < currentTestDataRow.Products.Count; productNumber++) {
                    Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + (productNumber * 2)).StringCellValue,
                        currentTestDataRow.Products[productNumber].Title);
                    Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + (productNumber * 2 + 1)).NumericCellValue,
                        Convert.ToDouble(currentTestDataRow.Products[productNumber].Amount));
                }

                for (var enumProp1Number = 0; enumProp1Number < currentTestDataRow.EnumProps1.Count; enumProp1Number++) {
                    Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + (enumProp1Number * 3)).StringCellValue,
                        currentTestDataRow.EnumProps1[enumProp1Number].Field1);
                    Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + (enumProp1Number * 3 + 1)).NumericCellValue,
                        Convert.ToDouble(currentTestDataRow.EnumProps1[enumProp1Number].Field2));
                    Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + (enumProp1Number * 3 + 2)).StringCellValue,
                        currentTestDataRow.EnumProps1[enumProp1Number].Field3.ToRussianString());
                }

                for (var enumProp2Number = 0; enumProp2Number < currentTestDataRow.EnumProps2.Count; enumProp2Number++) {
                    Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + testDataEnumProp1ColumnsQuantity + (enumProp2Number * 5)).StringCellValue,
                        currentTestDataRow.EnumProps2[enumProp2Number].Field4);
                    Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + testDataEnumProp1ColumnsQuantity + (enumProp2Number * 5 + 1)).NumericCellValue,
                        Convert.ToDouble(currentTestDataRow.EnumProps2[enumProp2Number].Field5));
                    Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + testDataEnumProp1ColumnsQuantity + (enumProp2Number * 5 + 2)).StringCellValue,
                        currentTestDataRow.EnumProps2[enumProp2Number].Field6.ToRussianString());
                    Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + testDataEnumProp1ColumnsQuantity + (enumProp2Number * 5 + 3)).StringCellValue,
                        currentTestDataRow.EnumProps2[enumProp2Number].Field7);
                    Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + testDataEnumProp1ColumnsQuantity + (enumProp2Number * 5 + 4)).NumericCellValue,
                        Convert.ToDouble(currentTestDataRow.EnumProps2[enumProp2Number].Field8));
                }
            }
		}
		[Test]
		public void ExcelStyleSimpleXlsxExport() {
		    const string filename = "TestSimpleStyle.xlsx";
            DeleteTestFile(filename);
			var simpleTestData = NotRelectionTestDataEntities.CreateSimpleTestData(true);
			var firstTestDataRow = simpleTestData.Models.FirstOrDefault();
			Assert.NotNull(firstTestDataRow);
			var style = new TableWriterStyle();
			var ms = TableWriterSimple.Write(new XlsxTableWriterSimple(style), simpleTestData.Models, simpleTestData.ConfigurationBuilder.Value);
            WriteToFile(ms, filename);
			var firstContact = firstTestDataRow.Contacts.FirstOrDefault();
			Assert.NotNull(firstContact);

			var firstContract = firstTestDataRow.Contracts.FirstOrDefault();
			Assert.NotNull(firstContract);

			var firstProduct = firstTestDataRow.Products.FirstOrDefault();
			Assert.NotNull(firstProduct);

			var firstEnumProp1 = firstTestDataRow.EnumProps1.FirstOrDefault();
			Assert.NotNull(firstEnumProp1);

			var firstEnumProp2 = firstTestDataRow.EnumProps2.FirstOrDefault();
			Assert.NotNull(firstEnumProp2);

			XSSFWorkbook xlsFile;
			using (var file = new FileStream(filename, FileMode.Open, FileAccess.Read)) {
				xlsFile = new XSSFWorkbook(file);
			}
			var sheet = xlsFile.GetSheetAt(0);

			short[] red = { 255, 0, 0 };
			short[] green = { 0, 255, 0 };
			short[] blue = { 0, 0, 255 };

			var rowNumber = 0;
			var row = sheet.GetRow(rowNumber);
			for (var cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
				CustomAssert.IsEqualFont(xlsFile, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Bold);
			}

			for (rowNumber = 1; rowNumber < sheet.LastRowNum; rowNumber++) {
				row = sheet.GetRow(rowNumber);
				for (var cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
					CustomAssert.IsEqualFont(xlsFile, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Normal);
				}
			}
		}
		[Test]
		public void ExcelComplexXlsxExport() {
		    const string filename = "TestComplex.xlsx";
            DeleteTestFile(filename);
			var simpleTestData = NotRelectionTestDataEntities.CreateSimpleTestData(true);
			var firstTestDataRow = simpleTestData.Models.FirstOrDefault();
			Assert.NotNull(firstTestDataRow);
			var style = new TableWriterStyle();
			var ms = TableWriterComplex.Write(new XlsxTableWriterComplex(style), simpleTestData.Models, simpleTestData.ConfigurationBuilder.Value);
            WriteToFile(ms, filename);
			XSSFWorkbook xlsFile;
			using (var file = new FileStream(filename, FileMode.Open, FileAccess.Read)) {
				xlsFile = new XSSFWorkbook(file);
			}
			var sheet = xlsFile.GetSheetAt(0);
			var parentColumnsNames = simpleTestData.ConfigurationBuilder.Value.ColumnsMap.Keys.ToList();

			var modelsQuantity = simpleTestData.Models.Count;
			var rowNumber = 0;
            for (var modelNumber = 0; modelNumber < modelsQuantity; modelNumber++) {
                var currentTestDataRow = simpleTestData.Models[modelNumber];
                var row = sheet.GetRow(rowNumber);
                var cell = row.GetCell(0);
                Assert.AreEqual(cell.StringCellValue, simpleTestData.ConfigurationBuilder.Value.Title);
                for (var cellNumber = 1; cellNumber < row.LastCellNum; cellNumber++) {
                    cell = row.GetCell(cellNumber);
                    Assert.AreEqual(cell.StringCellValue, parentColumnsNames[cellNumber - 1]);
                }
                rowNumber++;
                row = sheet.GetRow(rowNumber);

                Assert.AreEqual(row.GetCell(1).StringCellValue, currentTestDataRow.Title);
                Assert.AreEqual(row.GetCell(2).CellFormula, string.Format("=Date({0},{1},{2})", currentTestDataRow.RegistrationDate.Year, currentTestDataRow.RegistrationDate.Month, currentTestDataRow.RegistrationDate.Day));
                Assert.AreEqual(row.GetCell(3).StringCellValue, currentTestDataRow.Phone);
                Assert.AreEqual(row.GetCell(4).StringCellValue, currentTestDataRow.Inn);
                Assert.AreEqual(row.GetCell(5).StringCellValue, currentTestDataRow.Okato);
                Assert.AreEqual(row.GetCell(6).NumericCellValue, Convert.ToDouble(currentTestDataRow.Revenue));
                Assert.AreEqual(row.GetCell(7).NumericCellValue, Convert.ToDouble(currentTestDataRow.EmployeeCount));
                Assert.AreEqual(row.GetCell(8).StringCellValue, currentTestDataRow.IsActive.ToRussianString());
                Assert.AreEqual(row.GetCell(9).NumericCellValue, Convert.ToDouble(currentTestDataRow.Prop1));
                Assert.AreEqual(row.GetCell(10).StringCellValue, currentTestDataRow.Prop2);

                Assert.AreEqual(row.GetCell(11).StringCellValue, currentTestDataRow.Prop3.ToRussianString());
                Assert.AreEqual(row.GetCell(12).NumericCellValue, Convert.ToDouble(currentTestDataRow.Prop4));
                Assert.AreEqual(row.GetCell(13).StringCellValue, currentTestDataRow.Prop5);
                Assert.AreEqual(row.GetCell(14).StringCellValue, currentTestDataRow.Prop6.ToRussianString());
                Assert.AreEqual(row.GetCell(15).NumericCellValue, Convert.ToDouble(currentTestDataRow.Prop7));
                rowNumber++;
                row = sheet.GetRow(rowNumber);

                var childsQuantity = simpleTestData.ConfigurationBuilder.Value.ChildrenMap.Count;

                for (var childNumber = 0; childNumber < childsQuantity; childNumber++) {
                    switch (childNumber) {
                        case 0: {
                                var child = currentTestDataRow.Contacts;
                                TestChildHeader(row, childNumber, simpleTestData);
                                rowNumber++;
                                row = sheet.GetRow(rowNumber);
                                foreach (var childProperty in child) {
                                    Assert.AreEqual(row.GetCell(1).StringCellValue, childProperty.Title);
                                    Assert.AreEqual(row.GetCell(2).StringCellValue, childProperty.Email);
                                    rowNumber++;
                                    row = sheet.GetRow(rowNumber);
                                }
                                break;
                            }
                        case 1: {
                                var child = currentTestDataRow.Contracts;
                                TestChildHeader(row, childNumber, simpleTestData);
                                rowNumber++;
                                row = sheet.GetRow(rowNumber);
                                foreach (var childProperty in child) {
                                    Assert.AreEqual(row.GetCell(1).CellFormula, string.Format("=Date({0},{1},{2})", childProperty.BeginDate.Year, childProperty.BeginDate.Month, childProperty.BeginDate.Day));
                                    Assert.AreEqual(row.GetCell(2).CellFormula, string.Format("=Date({0},{1},{2})", childProperty.EndDate.Year, childProperty.EndDate.Month, childProperty.EndDate.Day));
                                    Assert.AreEqual(row.GetCell(3).StringCellValue, childProperty.Status.ToRussianString());
                                    rowNumber++;
                                    row = sheet.GetRow(rowNumber);
                                }
                                break;
                            }
                        case 2: {
                                var child = currentTestDataRow.Products;
                                TestChildHeader(row, childNumber, simpleTestData);
                                rowNumber++;
                                row = sheet.GetRow(rowNumber);
                                foreach (var childProperty in child) {
                                    Assert.AreEqual(row.GetCell(1).StringCellValue, childProperty.Title);
                                    Assert.AreEqual(row.GetCell(2).NumericCellValue, Convert.ToDouble(childProperty.Amount));
                                    rowNumber++;
                                    row = sheet.GetRow(rowNumber);
                                }
                                break;
                            }
                        case 3: {
                                var child = currentTestDataRow.EnumProps1;
                                TestChildHeader(row, childNumber, simpleTestData);
                                rowNumber++;
                                row = sheet.GetRow(rowNumber);
                                foreach (var childProperty in child) {
                                    Assert.AreEqual(row.GetCell(1).StringCellValue, childProperty.Field1);
                                    Assert.AreEqual(row.GetCell(2).NumericCellValue, Convert.ToDouble(childProperty.Field2));
                                    Assert.AreEqual(row.GetCell(3).StringCellValue, childProperty.Field3.ToRussianString());
                                    rowNumber++;
                                    row = sheet.GetRow(rowNumber);
                                }
                                break;
                            }
                        case 4: {
                                var child = currentTestDataRow.EnumProps2;
                                TestChildHeader(row, childNumber, simpleTestData);
                                rowNumber++;
                                row = sheet.GetRow(rowNumber);
                                foreach (var childProperty in child) {
                                    Assert.AreEqual(row.GetCell(1).StringCellValue, childProperty.Field4);
                                    Assert.AreEqual(row.GetCell(2).NumericCellValue, Convert.ToDouble(childProperty.Field5));
                                    Assert.AreEqual(row.GetCell(3).StringCellValue, childProperty.Field6.ToRussianString());
                                    Assert.AreEqual(row.GetCell(4).StringCellValue, childProperty.Field7);
                                    Assert.AreEqual(row.GetCell(5).NumericCellValue, Convert.ToDouble(childProperty.Field8));
                                    rowNumber++;
                                    row = sheet.GetRow(rowNumber);
                                }
                                break;
                            }
                    }
                }
            }
		}
		[Test]
		public void ExcelStyleComplexXlsxExport() {
		    const string filename = "TestComplexStyle.xlsx";
            DeleteTestFile(filename);
			var simpleTestData = NotRelectionTestDataEntities.CreateSimpleTestData(true);
			var firstTestDataRow = simpleTestData.Models.FirstOrDefault();
			Assert.NotNull(firstTestDataRow);
			var style = new TableWriterStyle();
			var ms = TableWriterComplex.Write(new XlsxTableWriterComplex(style), simpleTestData.Models, simpleTestData.ConfigurationBuilder.Value);
            WriteToFile(ms, filename);
			XSSFWorkbook xlsFile;
			using (var file = new FileStream(filename, FileMode.Open, FileAccess.Read)) {
				xlsFile = new XSSFWorkbook(file);
			}
			var sheet = xlsFile.GetSheetAt(0);

			short[] red = { 255, 0, 0 };
			short[] green = { 0, 255, 0 };
			short[] blue = { 0, 0, 255 };
			var modelsQuantity = simpleTestData.Models.Count;
			var rowNumber = 0;
			var childsQuantity = 0;
			for (var modelNumber = 0; modelNumber < modelsQuantity; modelNumber++) {
				childsQuantity += 7 + simpleTestData.Models[modelNumber].Contacts.Count + simpleTestData.Models[modelNumber].Contracts.Count + simpleTestData.Models[modelNumber].Products.Count + simpleTestData.Models[modelNumber].EnumProps1.Count + simpleTestData.Models[modelNumber].EnumProps2.Count;
				var row = sheet.GetRow(rowNumber);
				for (var cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
                    if (sheet.GetRow(rowNumber).GetCell(cellNumber) != null 
                        && !(sheet.GetRow(rowNumber).GetCell(cellNumber).CellType == CellType.String 
                        && string.IsNullOrWhiteSpace(sheet.GetRow(rowNumber).GetCell(cellNumber).StringCellValue))) {
				        CustomAssert.IsEqualFont(xlsFile, rowNumber, cellNumber, "Arial", 10, (short) FontBoldWeight.Bold);
				    }
				}
				rowNumber++;

				for (; rowNumber < childsQuantity; rowNumber++) {
					row = sheet.GetRow(rowNumber);
					for (var cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
                        if (sheet.GetRow(rowNumber).GetCell(cellNumber) != null
                        && !(sheet.GetRow(rowNumber).GetCell(cellNumber).CellType == CellType.String
                        && string.IsNullOrWhiteSpace(sheet.GetRow(rowNumber).GetCell(cellNumber).StringCellValue))) {
					        CustomAssert.IsEqualFont(xlsFile, rowNumber, cellNumber, "Arial", 10, (short) FontBoldWeight.Normal);
					    }
					}
				}
			}
		}
		[Test]
		public void ExcelReflectionSimpleXlsxExport() {
		    const string filename = "TestSimpleReflection.xlsx";
            DeleteTestFile(filename);
			var test = new List<ReflectionTestDataEntities>();
			var style = new TableWriterStyle();
			var rand = new Random();
			for (var i = 2; i < rand.Next(10, 30); i++) {
				test.Add(new ReflectionTestDataEntities());
			}
			var ms = ReflectionWriterSimple.Write(test, new XlsxTableWriterSimple(style), new CultureInfo("ru-Ru"));
            WriteToFile(ms, filename);
			ExcelReflectionSimpleExportTest(test, filename, true);

			ExcelStyleReflectionSimpleExportTest(test, filename, true);
		}
		[Test]
		public void ExcelReflectionComplexXlsxExport() {
		    const string filename = "TestReflectionComplex.xlsx";
            DeleteTestFile(filename);
			var test = new List<ReflectionTestDataEntities>();
			var style = new TableWriterStyle();
			var rand = new Random();
			for (var i = 2; i < rand.Next(10, 30); i++) {
				test.Add(new ReflectionTestDataEntities());
			}
			var ms = ReflectionWriterComplex.Write(test, new XlsxTableWriterComplex(style), new CultureInfo("ru-Ru"));
            WriteToFile(ms, filename);
			ExcelReflectionComplexExportTest(test, filename, true);

			ExcelStyleReflectionComplexExportTest(test, filename, true);
		}
		private static void ExcelReflectionSimpleExportTest<T>(List<T> models, string filename, bool isXlsx = false) {
			var firstModel = models.FirstOrDefault();
			Assert.NotNull(firstModel);
			var generalTypes = GeneralTypes;
			var nonEnumerableProperties = firstModel.GetType().GetProperties().Where(x => generalTypes.Contains(x.PropertyType)).ToList();
			var enumerableProperties = firstModel.GetType().GetProperties().Where(x => x.PropertyType.IsGenericType && x.PropertyType.GetGenericTypeDefinition() == typeof(List<>)).ToList();

			IWorkbook xlsFile;
			using (var file = new FileStream(filename, FileMode.Open, FileAccess.Read)) {
				if (isXlsx) {
					xlsFile = new XSSFWorkbook(file);
				} else {
					xlsFile = new HSSFWorkbook(file);
				}
			}
			var sheet = xlsFile.GetSheetAt(0);
			var rowNumber = 0;
			var row = sheet.GetRow(rowNumber);
			var cellNumber = 0;

			var enumerablePropertiesChildrensMaxQuantity = new int[enumerableProperties.Count()];
			for (var i = 0; i < enumerablePropertiesChildrensMaxQuantity.Count(); i++) {
				enumerablePropertiesChildrensMaxQuantity[i] = 0;
			}
			foreach (var model in models) {
				for (var i = 0; i < enumerableProperties.Count(); i++) {
					var nestedModels = enumerableProperties.ElementAt(i).GetValue(model);
					if (enumerablePropertiesChildrensMaxQuantity[i] < ((IList)nestedModels).Count) {
						enumerablePropertiesChildrensMaxQuantity[i] = ((IList)nestedModels).Count;
					}
				}
			}
			foreach (var property in nonEnumerableProperties) {
				var attribute = property.GetCustomAttribute<ExcelExportAttribute>();
				if (attribute != null && attribute.IsExportable) {
					Assert.AreEqual(row.GetCell(cellNumber).StringCellValue, attribute.PropertyName);
					cellNumber++;
				}
			}

			for (var i = 0; i < enumerableProperties.Count(); i++) {
				var property = enumerableProperties.ElementAt(i);
				var propertyType = property.PropertyType;
				var listType = propertyType.GetGenericArguments()[0];

				var props = listType.GetProperties().Where(x => generalTypes.Contains(x.PropertyType)).ToList();
				for (var j = 0; j < enumerablePropertiesChildrensMaxQuantity[i]; j++) {
					foreach (var prop in props) {
						var attribute = prop.GetCustomAttribute<ExcelExportAttribute>();
						Assert.AreEqual(row.GetCell(cellNumber).StringCellValue, attribute.PropertyName + (j + 1));
						cellNumber++;
					}
				}
			}
			var cultureInfo = new CultureInfo("ru-RU");
			rowNumber = 1;
			foreach (var model in models) {
				cellNumber = 0;
				row = sheet.GetRow(rowNumber);
				foreach (var nonEnumerableProperty in nonEnumerableProperties) {
					var attribute = nonEnumerableProperty.GetCustomAttribute<ExcelExportAttribute>();
					if (attribute != null && attribute.IsExportable) {
						var value = ConvertPropertyToExcelFormat(nonEnumerableProperty, model);
                        if (value is DateTime) Assert.AreEqual(DateUtil.GetJavaDate(row.GetCell(cellNumber).NumericCellValue).ToString(CultureInfo.InvariantCulture), actual: value.ToString(CultureInfo.InvariantCulture));
					    else {
					        if (row.GetCell(cellNumber).CellType == CellType.String) Assert.AreEqual(row.GetCell(cellNumber).StringCellValue, value);
					        if (row.GetCell(cellNumber).CellType == CellType.Numeric) Assert.AreEqual(row.GetCell(cellNumber).NumericCellValue, value);
					    }
					    cellNumber++;
					}
				}

				for (var i = 0; i < enumerableProperties.Count(); i++) {
					var property = enumerableProperties.ElementAt(i);
					var submodels = (IList)property.GetValue(model);

					var propertyType = property.PropertyType;
					var listType = propertyType.GetGenericArguments()[0];

					var listPeroperties = listType.GetProperties().Where(x => generalTypes.Contains(x.PropertyType)).ToList();

					foreach (var submodel in submodels) {
						foreach (var listProperty in listPeroperties) {
							var attribute = listProperty.GetCustomAttribute<ExcelExportAttribute>();
                            if (attribute != null && attribute.IsExportable) {
                                var value = ConvertPropertyToExcelFormat(listProperty, submodel);
                                if (value is DateTime) Assert.AreEqual(DateUtil.GetJavaDate(row.GetCell(cellNumber).NumericCellValue).ToString(CultureInfo.InvariantCulture), actual: value.ToString(CultureInfo.InvariantCulture));
                                else {
                                    if (row.GetCell(cellNumber).CellType == CellType.String) Assert.AreEqual(row.GetCell(cellNumber).StringCellValue, value);
                                    if (row.GetCell(cellNumber).CellType == CellType.Numeric) Assert.AreEqual(row.GetCell(cellNumber).NumericCellValue, value);
                                }
								cellNumber++;
							}
						}
					}
					cellNumber += (enumerablePropertiesChildrensMaxQuantity[i] - submodels.Count) * listPeroperties.Count();
				}
				rowNumber++;
			}
		}
		private static void ExcelStyleReflectionSimpleExportTest<T>(IEnumerable<T> models, string filename, bool isXlsx = false) {
			var firstModel = models.FirstOrDefault();
			Assert.NotNull(firstModel);

			IWorkbook xlsFile;
			using (var file = new FileStream(filename, FileMode.Open, FileAccess.Read)) {
				if (isXlsx) {
					xlsFile = new XSSFWorkbook(file);
				} else {
					xlsFile = new HSSFWorkbook(file);
				}
			}
			var sheet = xlsFile.GetSheetAt(0);

			var rowNumber = 0;
			var row = sheet.GetRow(rowNumber);
			for (var cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
				if (isXlsx) {
					CustomAssert.IsEqualFont((XSSFWorkbook)xlsFile, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Bold);
				} else {
					CustomAssert.IsEqualFont(xlsFile, sheet, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Bold);
				}
			}

			for (rowNumber = 1; rowNumber < sheet.LastRowNum; rowNumber++) {
				row = sheet.GetRow(rowNumber);
				for (var cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
					if (isXlsx) {
						CustomAssert.IsEqualFont((XSSFWorkbook)xlsFile, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Normal);
					} else {
						CustomAssert.IsEqualFont(xlsFile, sheet, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Normal);
					}
				}
			}
		}
		private static void ExcelReflectionComplexExportTest<T>(List<T> models, string filename, bool isXlsx = false) {
			var firstModel = models.FirstOrDefault();
			Assert.NotNull(firstModel);
			var generalTypes = GeneralTypes;
			var nonEnumerableProperties = firstModel.GetType().GetProperties().Where(x => generalTypes.Contains(x.PropertyType)).ToList();
			var enumerableProperties = firstModel.GetType().GetProperties().Where(x => x.PropertyType.IsGenericType && x.PropertyType.GetGenericTypeDefinition() == typeof(List<>)).ToList();

			IWorkbook xlsFile;
			using (var file = new FileStream(filename, FileMode.Open, FileAccess.Read)) {
				if (isXlsx) {
					xlsFile = new XSSFWorkbook(file);
				} else {
					xlsFile = new HSSFWorkbook(file);
				}
			}
			var sheet = xlsFile.GetSheetAt(0);
			var rowNumber = 0;
			var row = sheet.GetRow(rowNumber);
			var cellNumber = 0;
			var cultureInfo = new CultureInfo("ru-RU");
			foreach (var model in models) {
				Assert.AreEqual(row.GetCell(cellNumber).StringCellValue, model.GetType().GetCustomAttribute<ExcelExportClassNameAttribute>().Name);
				cellNumber++;

				foreach (var nonEnumerableProperty in nonEnumerableProperties) {
					var attribute = nonEnumerableProperty.GetCustomAttribute<ExcelExportAttribute>();
					if (attribute != null && attribute.IsExportable) {
						Assert.AreEqual(row.GetCell(cellNumber).StringCellValue, attribute.PropertyName);
						cellNumber++;
					}
				}
				rowNumber++;
				cellNumber = 1;

				row = sheet.GetRow(rowNumber);
				foreach (var nonEnumerableProperty in nonEnumerableProperties) {
					var attribute = nonEnumerableProperty.GetCustomAttribute<ExcelExportAttribute>();
					if (attribute != null && attribute.IsExportable) {
						var value = ConvertPropertyToExcelFormat(nonEnumerableProperty, model);
                        if (row.GetCell(cellNumber).CellType == CellType.String) Assert.AreEqual(row.GetCell(cellNumber).StringCellValue, value);
                        if (row.GetCell(cellNumber).CellType == CellType.Numeric) Assert.AreEqual(row.GetCell(cellNumber).NumericCellValue, value);
						cellNumber++;
					}
				}
				rowNumber++;
				cellNumber = 0;

				row = sheet.GetRow(rowNumber);
				foreach (var property in enumerableProperties) {
					var propertyType = property.PropertyType;
					var listType = propertyType.GetGenericArguments()[0];

					var props = listType.GetProperties().Where(x => generalTypes.Contains(x.PropertyType)).ToList();
					var submodels = (IList)property.GetValue(model);
					if (submodels.Count != 0) {
						var submodelName = submodels[0].GetType().GetCustomAttribute<ExcelExportClassNameAttribute>().Name;
						Assert.AreEqual(row.GetCell(cellNumber).StringCellValue, submodelName);
					}
					cellNumber++;
					foreach (var prop in props) {
						var attribute1 = prop.GetCustomAttribute<ExcelExportAttribute>();
						Assert.AreEqual(row.GetCell(cellNumber).StringCellValue, attribute1.PropertyName);
						cellNumber++;
					}
					rowNumber++;
					row = sheet.GetRow(rowNumber);
					if (((IList)property.GetValue(model)).Count != 0) {
						foreach (var submodel in submodels) {
							cellNumber = 1;
							foreach (var prop in props) {
								var attribute = prop.GetCustomAttribute<ExcelExportAttribute>();
								if (attribute != null && attribute.IsExportable) {
									var value = ConvertPropertyToExcelFormat(prop, submodel);
                                    if (row.GetCell(cellNumber).CellType == CellType.String) Assert.AreEqual(row.GetCell(cellNumber).StringCellValue, value);
                                    if (row.GetCell(cellNumber).CellType == CellType.Numeric) Assert.AreEqual(row.GetCell(cellNumber).NumericCellValue, value);
									cellNumber++;
								}
							}
							rowNumber++;
							row = sheet.GetRow(rowNumber);
						}
						cellNumber = 0;
					} else cellNumber = 0;
				}
			}
		}
		private static void ExcelStyleReflectionComplexExportTest<T>(List<T> models, string filename, bool isXlsx = false) {
			Assert.NotNull(models.FirstOrDefault());
			IWorkbook xlsFile;
			using (var file = new FileStream(filename, FileMode.Open, FileAccess.Read)) {
				if (isXlsx) {
					xlsFile = new XSSFWorkbook(file);
				} else {
					xlsFile = new HSSFWorkbook(file);
				}
			}
			var sheet = xlsFile.GetSheetAt(0);

			var rowNumber = 0;
			var row = sheet.GetRow(rowNumber);
			foreach (var model in models) {
				int cellNumber;
				for (cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
					if (isXlsx) {
                        if (sheet.GetRow(rowNumber).GetCell(cellNumber) != null 
                            && !(sheet.GetRow(rowNumber).GetCell(cellNumber).CellType == CellType.String 
                            && string.IsNullOrWhiteSpace(sheet.GetRow(rowNumber).GetCell(cellNumber).StringCellValue))) {
					        CustomAssert.IsEqualFont((XSSFWorkbook) xlsFile, rowNumber, cellNumber, "Arial", 10, (short) FontBoldWeight.Bold);
					    }
					} else {
                        if (sheet.GetRow(rowNumber).GetCell(cellNumber) != null
                            && !(sheet.GetRow(rowNumber).GetCell(cellNumber).CellType == CellType.String
                            && string.IsNullOrWhiteSpace(sheet.GetRow(rowNumber).GetCell(cellNumber).StringCellValue))) {
					        CustomAssert.IsEqualFont(xlsFile, sheet, rowNumber, cellNumber, "Arial", 10, (short) FontBoldWeight.Bold);
					    }
					}

				}
				rowNumber++;
				row = sheet.GetRow(rowNumber);
				var enumerableProperties = model.GetType().GetProperties().Where(x => x.PropertyType.IsGenericType && x.PropertyType.GetGenericTypeDefinition() == typeof(List<>)).ToList();
				for (cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
					if (isXlsx) {
                        if (sheet.GetRow(rowNumber).GetCell(cellNumber) != null
                            && !(sheet.GetRow(rowNumber).GetCell(cellNumber).CellType == CellType.String
                            && string.IsNullOrWhiteSpace(sheet.GetRow(rowNumber).GetCell(cellNumber).StringCellValue))) {
					        CustomAssert.IsEqualFont((XSSFWorkbook) xlsFile, rowNumber, cellNumber, "Arial", 10, (short) FontBoldWeight.Normal);
					    }
					} else {
                        if (sheet.GetRow(rowNumber).GetCell(cellNumber) != null
                            && !(sheet.GetRow(rowNumber).GetCell(cellNumber).CellType == CellType.String
                            && string.IsNullOrWhiteSpace(sheet.GetRow(rowNumber).GetCell(cellNumber).StringCellValue))) {
					        CustomAssert.IsEqualFont(xlsFile, sheet, rowNumber, cellNumber, "Arial", 10, (short) FontBoldWeight.Normal);
					    }
					}
				}
				rowNumber++;
				row = sheet.GetRow(rowNumber);
				foreach (var enumerableProperty in enumerableProperties) {
					var childs = (IList)enumerableProperty.GetValue(model);
					for (cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
						if (isXlsx) {
                            if (sheet.GetRow(rowNumber).GetCell(cellNumber) != null
                            && !(sheet.GetRow(rowNumber).GetCell(cellNumber).CellType == CellType.String
                            && string.IsNullOrWhiteSpace(sheet.GetRow(rowNumber).GetCell(cellNumber).StringCellValue))) {
						        CustomAssert.IsEqualFont((XSSFWorkbook) xlsFile, rowNumber, cellNumber, "Arial", 10, (short) FontBoldWeight.Normal);
						    }
						} else {
                            if (sheet.GetRow(rowNumber).GetCell(cellNumber) != null
                            && !(sheet.GetRow(rowNumber).GetCell(cellNumber).CellType == CellType.String
                            && string.IsNullOrWhiteSpace(sheet.GetRow(rowNumber).GetCell(cellNumber).StringCellValue))) {
						        CustomAssert.IsEqualFont(xlsFile, sheet, rowNumber, cellNumber, "Arial", 10, (short) FontBoldWeight.Normal);
						    }
						}
					}
					rowNumber++;

					var rowNumberStart = rowNumber;
					for (rowNumber = rowNumberStart; rowNumber < rowNumberStart + childs.Count; rowNumber++) {
						row = sheet.GetRow(rowNumber);
						for (cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
							if (isXlsx) {
                                if (sheet.GetRow(rowNumber).GetCell(cellNumber) != null
                            && !(sheet.GetRow(rowNumber).GetCell(cellNumber).CellType == CellType.String
                            && string.IsNullOrWhiteSpace(sheet.GetRow(rowNumber).GetCell(cellNumber).StringCellValue))) {
							        CustomAssert.IsEqualFont((XSSFWorkbook) xlsFile, rowNumber, cellNumber, "Arial", 10, (short) FontBoldWeight.Normal);
							    }
							} else {
                                if (sheet.GetRow(rowNumber).GetCell(cellNumber) != null
                            && !(sheet.GetRow(rowNumber).GetCell(cellNumber).CellType == CellType.String
                            && string.IsNullOrWhiteSpace(sheet.GetRow(rowNumber).GetCell(cellNumber).StringCellValue))) {
							        CustomAssert.IsEqualFont(xlsFile, sheet, rowNumber, cellNumber, "Arial", 10, (short) FontBoldWeight.Normal);
							    }
							}
						}
					}
					row = sheet.GetRow(rowNumber);
				}
			}
		}
        private static dynamic ConvertPropertyToExcelFormat<T>(PropertyInfo nonEnumerableProperty, T model) {
            var propertyTypeName = nonEnumerableProperty.PropertyType.Name;
            dynamic value = String.Empty;
            if (nonEnumerableProperty.GetValue(model) != null) {
                switch (propertyTypeName) {
                    case "String":
                        value = nonEnumerableProperty.GetValue(model).ToString();
                        break;
                    case "DateTime":
                        value = nonEnumerableProperty.GetValue(model);
                        break;
                    case "Decimal":
                        value = Convert.ToDouble(nonEnumerableProperty.GetValue(model));
                        break;
                    case "Int32":
                        value = Convert.ToDouble(nonEnumerableProperty.GetValue(model));
                        break;
                    case "Boolean":
                        value = ((bool)nonEnumerableProperty.GetValue(model)) ? "Да" : "Нет";
                        break;
                }
            }
            return value;
        }
		private static List<string> TestChildHeader(IRow row, int childNumber, NotRelectionTestDataEntities.TestData simpleTestData) {
			var childName = simpleTestData.ConfigurationBuilder.Value.ChildrenMap[childNumber].Title;
			var childColumnsNames = simpleTestData.ConfigurationBuilder.Value.ChildrenMap[childNumber].ColumnsMap.Keys.ToList();

			Assert.AreEqual(row.GetCell(0).StringCellValue, childName);
			for (var cellNumber = 1; cellNumber < row.LastCellNum; cellNumber++) {
				var numberProperty = cellNumber - 1;
				Assert.AreEqual(row.GetCell(cellNumber).StringCellValue, childColumnsNames[numberProperty]);
			}

			return childColumnsNames;
		}
		private static void WriteToFile(MemoryStream ms, string filename) {
			using (var file = new FileStream(filename, FileMode.Create, FileAccess.Write)) {
				var bytes = new byte[ms.Length];
				ms.Read(bytes, 0, (int)ms.Length);
				file.Write(bytes, 0, bytes.Length);
				ms.Close();
			}
		}

	    private static void DeleteTestFile(string filename) {
	        if (File.Exists(filename)) {
	            File.Delete(filename);
	        }
	    }
		private enum FontBoldWeight {
			Normal = 400,
			Bold = 700,
		}
		private static List<Type> GeneralTypes {
			get {
				var generalTypes = new List<Type> {
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