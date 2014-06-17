using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using Extensions;
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
		[Test]
		public void ExcelSimpleExportRowOverflow() {
			NotRelectionTestDataEntities.TestData simpleTestData = NotRelectionTestDataEntities.CreateSimpleTestRowOverflowData();
			NotRelectionTestDataEntities.ClientExampleModel firstTestDataRow = simpleTestData.Models.FirstOrDefault();
			Assert.NotNull(firstTestDataRow);
			MemoryStream memoryStream = TableWriterSimple.Write(new XlsTableWriterSimple(), simpleTestData.Models, simpleTestData.ConfigurationBuilder.Value);
			WriteToFile(memoryStream, "TestSimpleOverflow.xls");
			NotRelectionTestDataEntities.ClientExampleModel.Contact firstContact = firstTestDataRow.Contacts.FirstOrDefault();
			Assert.NotNull(firstContact);

			NotRelectionTestDataEntities.ClientExampleModel.Contract firstContract = firstTestDataRow.Contracts.FirstOrDefault();
			Assert.NotNull(firstContract);

			NotRelectionTestDataEntities.ClientExampleModel.Product firstProduct = firstTestDataRow.Products.FirstOrDefault();
			Assert.NotNull(firstProduct);

			NotRelectionTestDataEntities.ClientExampleModel.EnumProp1 firstEnumProp1 = firstTestDataRow.EnumProps1.FirstOrDefault();
			Assert.NotNull(firstEnumProp1);

			NotRelectionTestDataEntities.ClientExampleModel.EnumProp2 firstEnumProp2 = firstTestDataRow.EnumProps2.FirstOrDefault();
			Assert.NotNull(firstEnumProp2);

			int parentColumnsQuantity = simpleTestData.ConfigurationBuilder.Value.ColumnsMap.Count;
			int contactsFieldsQuantity = firstContact.GetType().GetProperties().Count();
			int contractsFieldsQuantity = firstContract.GetType().GetProperties().Count();
			int productsFieldsQuantity = firstProduct.GetType().GetProperties().Count();
			int enumProp1FieldsQuantity = firstEnumProp1.GetType().GetProperties().Count();

			int testDataContactColumnsQuantity = simpleTestData.Models.Max(x => x.Contacts.Count) * contactsFieldsQuantity;
			int testDataContractColumnsQuantity = simpleTestData.Models.Max(x => x.Contracts.Count) * contractsFieldsQuantity;
			int testDataProductColumnsQuantity = simpleTestData.Models.Max(x => x.Products.Count) * productsFieldsQuantity;
			int testDataEnumProp1ColumnsQuantity = simpleTestData.Models.Max(x => x.EnumProps1.Count) * enumProp1FieldsQuantity;

			List<string> parentColumnsNames = simpleTestData.ConfigurationBuilder.Value.ColumnsMap.Keys.ToList();
			List<List<string>> childsColumnsNames = simpleTestData.ConfigurationBuilder.Value.ChildrenMap.Select(childs => childs.ColumnsMap.Keys.ToList()).ToList();
			HSSFWorkbook xlsFile;
			using (FileStream file = new FileStream("TestSimpleOverflow.xls", FileMode.Open, FileAccess.Read)) {
				xlsFile = new HSSFWorkbook(file);
			}

			List<string> header = parentColumnsNames.ToList();
			for (int childNumber = 0; childNumber < childsColumnsNames.Count; childNumber++) {
				List<string> child = childsColumnsNames[childNumber];
				int childColumnsQuantity = 0;
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
				for (int index = 1; index <= childColumnsQuantity; index++) {
					header.AddRange(child.Select(childPropertyName => childPropertyName + index));
				}
			}
			int modelNumber = 0;
			for (int sheetNumber = 0; sheetNumber < xlsFile.NumberOfSheets; sheetNumber++) {
				ISheet sheet = xlsFile.GetSheetAt(sheetNumber);
				IRow row = sheet.GetRow(0);
				for (int columnNumber = 0; columnNumber < row.LastCellNum; columnNumber++) {
					Assert.AreEqual(row.GetCell(columnNumber).StringCellValue, header[columnNumber]);
				}
				for (int rowNumber = 1; rowNumber < sheet.LastRowNum; rowNumber++) {
					row = sheet.GetRow(rowNumber);
					Assert.NotNull(row);
					NotRelectionTestDataEntities.ClientExampleModel currentTestDataRow = simpleTestData.Models[modelNumber];
					Assert.AreEqual(row.GetCell(0).StringCellValue, currentTestDataRow.Title);
					Assert.AreEqual(row.GetCell(1).StringCellValue, currentTestDataRow.RegistrationDate.ToRussianFullString());
					Assert.AreEqual(row.GetCell(2).StringCellValue, currentTestDataRow.Phone);
					Assert.AreEqual(row.GetCell(3).StringCellValue, currentTestDataRow.Inn);
					Assert.AreEqual(row.GetCell(4).StringCellValue, currentTestDataRow.Okato);

					Assert.AreEqual(row.GetCell(5).StringCellValue, string.Format(new CultureInfo("ru-RU"), "{0:C}", currentTestDataRow.Revenue));
					Assert.AreEqual(row.GetCell(6).StringCellValue, currentTestDataRow.EmployeeCount.ToString(CultureInfo.InvariantCulture));
					Assert.AreEqual(row.GetCell(7).StringCellValue, currentTestDataRow.IsActive.ToRussianString());
					Assert.AreEqual(row.GetCell(8).StringCellValue, currentTestDataRow.Prop1.ToString(CultureInfo.InvariantCulture));
					Assert.AreEqual(row.GetCell(9).StringCellValue, currentTestDataRow.Prop2);

					Assert.AreEqual(row.GetCell(10).StringCellValue, currentTestDataRow.Prop3.ToRussianString());
					Assert.AreEqual(row.GetCell(11).StringCellValue, currentTestDataRow.Prop4.ToString(CultureInfo.InvariantCulture));
					Assert.AreEqual(row.GetCell(12).StringCellValue, currentTestDataRow.Prop5);
					Assert.AreEqual(row.GetCell(13).StringCellValue, currentTestDataRow.Prop6.ToRussianString());
					Assert.AreEqual(row.GetCell(14).StringCellValue, currentTestDataRow.Prop7.ToString(CultureInfo.InvariantCulture));

					for (int contactNumber = 0; contactNumber < currentTestDataRow.Contacts.Count; contactNumber++) {
						NotRelectionTestDataEntities.ClientExampleModel.Contact currentContactRow = currentTestDataRow.Contacts[contactNumber];
						Assert.AreEqual(row.GetCell(parentColumnsQuantity + (contactNumber * contactsFieldsQuantity)).StringCellValue, currentContactRow.Title);
						Assert.AreEqual(row.GetCell(parentColumnsQuantity + (contactNumber * contactsFieldsQuantity + 1)).StringCellValue, currentContactRow.Email);
					}

					for (int contractNumber = 0; contractNumber < currentTestDataRow.Contracts.Count; contractNumber++) {
						Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + (contractNumber * contractsFieldsQuantity)).StringCellValue,
							currentTestDataRow.Contracts[contractNumber].BeginDate.ToRussianFullString());
						Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + (contractNumber * contractsFieldsQuantity + 1)).StringCellValue,
							currentTestDataRow.Contracts[contractNumber].EndDate.ToRussianFullString());
						Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + (contractNumber * contractsFieldsQuantity + 2)).StringCellValue,
							currentTestDataRow.Contracts[contractNumber].Status.ToRussianString());
					}

					for (int productNumber = 0; productNumber < currentTestDataRow.Products.Count; productNumber++) {
						Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + (productNumber * 2)).StringCellValue,
							currentTestDataRow.Products[productNumber].Title);
						Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + (productNumber * 2 + 1)).StringCellValue,
							currentTestDataRow.Products[productNumber].Amount.ToString(CultureInfo.InvariantCulture));
					}

					for (int enumProp1Number = 0; enumProp1Number < currentTestDataRow.EnumProps1.Count; enumProp1Number++) {
						Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + (enumProp1Number * 3)).StringCellValue,
							currentTestDataRow.EnumProps1[enumProp1Number].Field1);
						Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + (enumProp1Number * 3 + 1)).StringCellValue,
							currentTestDataRow.EnumProps1[enumProp1Number].Field2.ToString(CultureInfo.InvariantCulture));
						Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + (enumProp1Number * 3 + 2)).StringCellValue,
							currentTestDataRow.EnumProps1[enumProp1Number].Field3.ToRussianString());
					}

					for (int enumProp2Number = 0; enumProp2Number < currentTestDataRow.EnumProps2.Count; enumProp2Number++) {
						Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + testDataEnumProp1ColumnsQuantity + (enumProp2Number * 5)).StringCellValue,
							currentTestDataRow.EnumProps2[enumProp2Number].Field4);
						Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + testDataEnumProp1ColumnsQuantity + (enumProp2Number * 5 + 1)).StringCellValue,
							currentTestDataRow.EnumProps2[enumProp2Number].Field5.ToString(CultureInfo.InvariantCulture));
						Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + testDataEnumProp1ColumnsQuantity + (enumProp2Number * 5 + 2)).StringCellValue,
							currentTestDataRow.EnumProps2[enumProp2Number].Field6.ToRussianString());
						Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + testDataEnumProp1ColumnsQuantity + (enumProp2Number * 5 + 3)).StringCellValue,
							currentTestDataRow.EnumProps2[enumProp2Number].Field7);
						Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + testDataEnumProp1ColumnsQuantity + (enumProp2Number * 5 + 4)).StringCellValue,
							currentTestDataRow.EnumProps2[enumProp2Number].Field8.ToString(CultureInfo.InvariantCulture));
					}
					modelNumber++;
				}
				modelNumber++;
			}
		}
		[Test]
		public void ExcelComplexExportRowOverflow() {
			NotRelectionTestDataEntities.TestData simpleTestData = NotRelectionTestDataEntities.CreateSimpleTestRowOverflowData();
			NotRelectionTestDataEntities.ClientExampleModel firstTestDataRow = simpleTestData.Models.FirstOrDefault();
			Assert.NotNull(firstTestDataRow);
			TableWriterStyle style = new TableWriterStyle();
			MemoryStream memoryStream = TableWriterComplex.Write(new XlsTableWriterComplex(style), simpleTestData.Models, simpleTestData.ConfigurationBuilder.Value);
			WriteToFile(memoryStream, "TestComplexOverflow.xls");

			HSSFWorkbook xlsFile;
			using (FileStream file = new FileStream("TestComplexOverflow.xls", FileMode.Open, FileAccess.Read)) {
				xlsFile = new HSSFWorkbook(file);
			}
			List<string> parentColumnsNames = simpleTestData.ConfigurationBuilder.Value.ColumnsMap.Keys.ToList();
			int sheetNumber = 0;
			int modelsQuantity = simpleTestData.Models.Count;
			int rowNumber = 0;

			ISheet sheet = xlsFile.GetSheetAt(sheetNumber);

			for (int modelNumber = 0; modelNumber < modelsQuantity; modelNumber++) {
				List<string> lastParentHeader = new List<string>();
				List<string> lastChildHeader = new List<string>();
				NotRelectionTestDataEntities.ClientExampleModel currentTestDataRow = simpleTestData.Models[modelNumber];
				IRow row = sheet.GetRow(rowNumber);
				ICell cell = row.GetCell(0);
				Assert.AreEqual(cell.StringCellValue, simpleTestData.ConfigurationBuilder.Value.Title);
				for (int cellNumber = 1; cellNumber < row.LastCellNum; cellNumber++) {
					cell = row.GetCell(cellNumber);
					Assert.AreEqual(cell.StringCellValue, parentColumnsNames[cellNumber - 1]);
					lastParentHeader.Add(parentColumnsNames[cellNumber - 1]);
				}
				RowNumberIncrement(lastParentHeader, lastChildHeader, xlsFile, ref rowNumber, ref sheet, ref sheetNumber);
				row = sheet.GetRow(rowNumber);

				Assert.AreEqual(row.GetCell(1).StringCellValue, currentTestDataRow.Title);
				Assert.AreEqual(row.GetCell(2).StringCellValue, currentTestDataRow.RegistrationDate.ToRussianFullString());
				Assert.AreEqual(row.GetCell(3).StringCellValue, currentTestDataRow.Phone);
				Assert.AreEqual(row.GetCell(4).StringCellValue, currentTestDataRow.Inn);
				Assert.AreEqual(row.GetCell(5).StringCellValue, currentTestDataRow.Okato);
				Assert.AreEqual(row.GetCell(6).StringCellValue, string.Format(new CultureInfo("ru-RU"), "{0:C}", currentTestDataRow.Revenue));
				Assert.AreEqual(row.GetCell(7).StringCellValue, currentTestDataRow.EmployeeCount.ToString(CultureInfo.InvariantCulture));
				Assert.AreEqual(row.GetCell(8).StringCellValue, currentTestDataRow.IsActive.ToRussianString());
				Assert.AreEqual(row.GetCell(9).StringCellValue, currentTestDataRow.Prop1.ToString(CultureInfo.InvariantCulture));
				Assert.AreEqual(row.GetCell(10).StringCellValue, currentTestDataRow.Prop2);
				Assert.AreEqual(row.GetCell(11).StringCellValue, currentTestDataRow.Prop3.ToRussianString());
				Assert.AreEqual(row.GetCell(12).StringCellValue, currentTestDataRow.Prop4.ToString(CultureInfo.InvariantCulture));
				Assert.AreEqual(row.GetCell(13).StringCellValue, currentTestDataRow.Prop5);
				Assert.AreEqual(row.GetCell(14).StringCellValue, currentTestDataRow.Prop6.ToRussianString());
				Assert.AreEqual(row.GetCell(15).StringCellValue, currentTestDataRow.Prop7.ToString(CultureInfo.InvariantCulture));
				RowNumberIncrement(lastParentHeader, null, xlsFile, ref rowNumber, ref sheet, ref sheetNumber);
				row = sheet.GetRow(rowNumber);

				int childsQuantity = simpleTestData.ConfigurationBuilder.Value.ChildrenMap.Count;

				for (int childNumber = 0; childNumber < childsQuantity; childNumber++) {
					switch (childNumber) {
						case 0: {
								List<NotRelectionTestDataEntities.ClientExampleModel.Contact> child = currentTestDataRow.Contacts;
								lastChildHeader = TestChildHeader(row, childNumber, simpleTestData);
								RowNumberIncrement(lastParentHeader, lastChildHeader, xlsFile, ref rowNumber, ref sheet, ref sheetNumber);
								row = sheet.GetRow(rowNumber);
								foreach (NotRelectionTestDataEntities.ClientExampleModel.Contact childProperty in child) {
									Assert.AreEqual(row.GetCell(1).StringCellValue, childProperty.Title);
									Assert.AreEqual(row.GetCell(2).StringCellValue, childProperty.Email);
									RowNumberIncrement(lastParentHeader, lastChildHeader, xlsFile, ref rowNumber, ref sheet, ref sheetNumber); 
									row = sheet.GetRow(rowNumber);
								}
								break;
							}
						case 1: {
								List<NotRelectionTestDataEntities.ClientExampleModel.Contract> child = currentTestDataRow.Contracts;
								lastChildHeader = TestChildHeader(row, childNumber, simpleTestData);
								RowNumberIncrement(lastParentHeader, lastChildHeader, xlsFile, ref rowNumber, ref sheet, ref sheetNumber);
								row = sheet.GetRow(rowNumber);
								foreach (NotRelectionTestDataEntities.ClientExampleModel.Contract childProperty in child) {
									Assert.AreEqual(row.GetCell(1).StringCellValue, childProperty.BeginDate.ToRussianFullString());
									Assert.AreEqual(row.GetCell(2).StringCellValue, childProperty.EndDate.ToRussianFullString());
									Assert.AreEqual(row.GetCell(3).StringCellValue, childProperty.Status.ToRussianString());
									RowNumberIncrement(lastParentHeader, lastChildHeader, xlsFile, ref rowNumber, ref sheet, ref sheetNumber); 
									row = sheet.GetRow(rowNumber);
								}
								break;
							}
						case 2: {
								List<NotRelectionTestDataEntities.ClientExampleModel.Product> child = currentTestDataRow.Products;
								lastChildHeader = TestChildHeader(row, childNumber, simpleTestData);
								RowNumberIncrement(lastParentHeader, lastChildHeader, xlsFile, ref rowNumber, ref sheet, ref sheetNumber);
								row = sheet.GetRow(rowNumber);
								foreach (NotRelectionTestDataEntities.ClientExampleModel.Product childProperty in child) {
									Assert.AreEqual(row.GetCell(1).StringCellValue, childProperty.Title);
									Assert.AreEqual(row.GetCell(2).StringCellValue, childProperty.Amount.ToString(CultureInfo.InvariantCulture));
									RowNumberIncrement(lastParentHeader, lastChildHeader, xlsFile, ref rowNumber, ref sheet, ref sheetNumber); 
									row = sheet.GetRow(rowNumber);
								}
								break;
							}
						case 3: {
								List<NotRelectionTestDataEntities.ClientExampleModel.EnumProp1> child = currentTestDataRow.EnumProps1;
								lastChildHeader = TestChildHeader(row, childNumber, simpleTestData);
								RowNumberIncrement(lastParentHeader, lastChildHeader, xlsFile, ref rowNumber, ref sheet, ref sheetNumber);
								row = sheet.GetRow(rowNumber);
								foreach (NotRelectionTestDataEntities.ClientExampleModel.EnumProp1 childProperty in child) {
									Assert.AreEqual(row.GetCell(1).StringCellValue, childProperty.Field1);
									Assert.AreEqual(row.GetCell(2).StringCellValue, childProperty.Field2.ToString(CultureInfo.InvariantCulture));
									Assert.AreEqual(row.GetCell(3).StringCellValue, childProperty.Field3.ToRussianString());
									RowNumberIncrement(lastParentHeader, lastChildHeader, xlsFile, ref rowNumber, ref sheet, ref sheetNumber); 
									row = sheet.GetRow(rowNumber);
								}
								break;
							}
						case 4: {
								List<NotRelectionTestDataEntities.ClientExampleModel.EnumProp2> child = currentTestDataRow.EnumProps2;
								lastChildHeader = TestChildHeader(row, childNumber, simpleTestData);
								RowNumberIncrement(lastParentHeader, lastChildHeader, xlsFile, ref rowNumber, ref sheet, ref sheetNumber);
								row = sheet.GetRow(rowNumber);
								foreach (NotRelectionTestDataEntities.ClientExampleModel.EnumProp2 childProperty in child) {
									Assert.AreEqual(row.GetCell(1).StringCellValue, childProperty.Field4);
									Assert.AreEqual(row.GetCell(2).StringCellValue, childProperty.Field5.ToString(CultureInfo.InvariantCulture));
									Assert.AreEqual(row.GetCell(3).StringCellValue, childProperty.Field6.ToRussianString());
									Assert.AreEqual(row.GetCell(4).StringCellValue, childProperty.Field7);
									Assert.AreEqual(row.GetCell(5).StringCellValue, childProperty.Field8.ToString(CultureInfo.InvariantCulture));
									RowNumberIncrement(lastParentHeader, lastChildHeader, xlsFile, ref rowNumber, ref sheet, ref sheetNumber); 
									row = sheet.GetRow(rowNumber);
								}
								break;
							}
					}
				}
			}
		}
		private static void RowNumberIncrement(IReadOnlyList<string> lastParentHeader, List<string> lastChildHeader, IWorkbook workbook, ref int rowNumber, ref ISheet sheet, ref int sheetNumber) {
			if (rowNumber >= 65535) {
				sheetNumber++;
				sheet = workbook.GetSheetAt(sheetNumber);
				IRow row = sheet.GetRow(0);
				for (int columnNumber = 1; columnNumber < row.LastCellNum; columnNumber++) {
					Assert.AreEqual(row.GetCell(columnNumber).StringCellValue, lastParentHeader[columnNumber - 1]);
				}
				row = sheet.GetRow(1);
				if (lastChildHeader != null) {
					for (int columnNumber = 1; columnNumber < row.LastCellNum; columnNumber++) {
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
			NotRelectionTestDataEntities.TestData simpleTestData = NotRelectionTestDataEntities.CreateSimpleTestData();
			NotRelectionTestDataEntities.ClientExampleModel firstTestDataRow = simpleTestData.Models.FirstOrDefault();
			Assert.NotNull(firstTestDataRow);
			MemoryStream memoryStream = TableWriterSimple.Write(new XlsTableWriterSimple(), simpleTestData.Models, simpleTestData.ConfigurationBuilder.Value);
			WriteToFile(memoryStream, "TestSimple.xls");

			NotRelectionTestDataEntities.ClientExampleModel.Contact firstContact = firstTestDataRow.Contacts.FirstOrDefault();
			Assert.NotNull(firstContact);

			NotRelectionTestDataEntities.ClientExampleModel.Contract firstContract = firstTestDataRow.Contracts.FirstOrDefault();
			Assert.NotNull(firstContract);

			NotRelectionTestDataEntities.ClientExampleModel.Product firstProduct = firstTestDataRow.Products.FirstOrDefault();
			Assert.NotNull(firstProduct);

			NotRelectionTestDataEntities.ClientExampleModel.EnumProp1 firstEnumProp1 = firstTestDataRow.EnumProps1.FirstOrDefault();
			Assert.NotNull(firstEnumProp1);

			NotRelectionTestDataEntities.ClientExampleModel.EnumProp2 firstEnumProp2 = firstTestDataRow.EnumProps2.FirstOrDefault();
			Assert.NotNull(firstEnumProp2);

			int parentColumnsQuantity = simpleTestData.ConfigurationBuilder.Value.ColumnsMap.Count;
			int contactsFieldsQuantity = firstContact.GetType().GetProperties().Count();
			int contractsFieldsQuantity = firstContract.GetType().GetProperties().Count();
			int productsFieldsQuantity = firstProduct.GetType().GetProperties().Count();
			int enumProp1FieldsQuantity = firstEnumProp1.GetType().GetProperties().Count();

			int testDataContactColumnsQuantity = simpleTestData.Models.Max(x => x.Contacts.Count) * contactsFieldsQuantity;
			int testDataContractColumnsQuantity = simpleTestData.Models.Max(x => x.Contracts.Count) * contractsFieldsQuantity;
			int testDataProductColumnsQuantity = simpleTestData.Models.Max(x => x.Products.Count) * productsFieldsQuantity;
			int testDataEnumProp1ColumnsQuantity = simpleTestData.Models.Max(x => x.EnumProps1.Count) * enumProp1FieldsQuantity;

			List<string> parentColumnsNames = simpleTestData.ConfigurationBuilder.Value.ColumnsMap.Keys.ToList();
			List<List<string>> childsColumnsNames = simpleTestData.ConfigurationBuilder.Value.ChildrenMap.Select(childs => childs.ColumnsMap.Keys.ToList()).ToList();
			HSSFWorkbook xlsFile;
			using (FileStream file = new FileStream("TestSimple.xls", FileMode.Open, FileAccess.Read)) {
				xlsFile = new HSSFWorkbook(file);
			}
			ISheet sheet = xlsFile.GetSheetAt(0);
			int rowNumber = 0;
			IRow row = sheet.GetRow(rowNumber);

			int columnNumber;
			for (columnNumber = 0; columnNumber < parentColumnsNames.Count; columnNumber++) {
				Assert.AreEqual(row.GetCell(columnNumber).StringCellValue, parentColumnsNames[columnNumber]);
			}

			columnNumber = parentColumnsQuantity;
			for (int childNumber = 0; childNumber < childsColumnsNames.Count; childNumber++) {
				List<string> child = childsColumnsNames[childNumber];
				int childColumnsQuantity = 0;
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
				for (int index = 1; index <= childColumnsQuantity; index++) {
					foreach (string childPropertyName in child) {
						Assert.AreEqual(row.GetCell(columnNumber).StringCellValue, childPropertyName + index);
						columnNumber++;
					}
				}
			}

			for (rowNumber = 1; rowNumber <= sheet.LastRowNum; rowNumber++) {
				row = sheet.GetRow(rowNumber);
				Assert.NotNull(row);
				NotRelectionTestDataEntities.ClientExampleModel currentTestDataRow = simpleTestData.Models[rowNumber - 1];
				Assert.AreEqual(row.GetCell(0).StringCellValue, currentTestDataRow.Title);
				Assert.AreEqual(row.GetCell(1).StringCellValue, currentTestDataRow.RegistrationDate.ToRussianFullString());
				Assert.AreEqual(row.GetCell(2).StringCellValue, currentTestDataRow.Phone);
				Assert.AreEqual(row.GetCell(3).StringCellValue, currentTestDataRow.Inn);
				Assert.AreEqual(row.GetCell(4).StringCellValue, currentTestDataRow.Okato);

				Assert.AreEqual(row.GetCell(5).StringCellValue, string.Format(new CultureInfo("ru-RU"), "{0:C}", currentTestDataRow.Revenue));
				Assert.AreEqual(row.GetCell(6).StringCellValue, currentTestDataRow.EmployeeCount.ToString(CultureInfo.InvariantCulture));
				Assert.AreEqual(row.GetCell(7).StringCellValue, currentTestDataRow.IsActive.ToRussianString());
				Assert.AreEqual(row.GetCell(8).StringCellValue, currentTestDataRow.Prop1.ToString(CultureInfo.InvariantCulture));
				Assert.AreEqual(row.GetCell(9).StringCellValue, currentTestDataRow.Prop2);

				Assert.AreEqual(row.GetCell(10).StringCellValue, currentTestDataRow.Prop3.ToRussianString());
				Assert.AreEqual(row.GetCell(11).StringCellValue, currentTestDataRow.Prop4.ToString(CultureInfo.InvariantCulture));
				Assert.AreEqual(row.GetCell(12).StringCellValue, currentTestDataRow.Prop5);
				Assert.AreEqual(row.GetCell(13).StringCellValue, currentTestDataRow.Prop6.ToRussianString());
				Assert.AreEqual(row.GetCell(14).StringCellValue, currentTestDataRow.Prop7.ToString(CultureInfo.InvariantCulture));

				for (int contactNumber = 0; contactNumber < currentTestDataRow.Contacts.Count; contactNumber++) {
					NotRelectionTestDataEntities.ClientExampleModel.Contact currentContactRow = currentTestDataRow.Contacts[contactNumber];
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + (contactNumber * contactsFieldsQuantity)).StringCellValue, currentContactRow.Title);
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + (contactNumber * contactsFieldsQuantity + 1)).StringCellValue, currentContactRow.Email);
				}

				for (int contractNumber = 0; contractNumber < currentTestDataRow.Contracts.Count; contractNumber++) {
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + (contractNumber * contractsFieldsQuantity)).StringCellValue,
						currentTestDataRow.Contracts[contractNumber].BeginDate.ToRussianFullString());
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + (contractNumber * contractsFieldsQuantity + 1)).StringCellValue,
						currentTestDataRow.Contracts[contractNumber].EndDate.ToRussianFullString());
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + (contractNumber * contractsFieldsQuantity + 2)).StringCellValue,
						currentTestDataRow.Contracts[contractNumber].Status.ToRussianString());
				}

				for (int productNumber = 0; productNumber < currentTestDataRow.Products.Count; productNumber++) {
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + (productNumber * 2)).StringCellValue,
						currentTestDataRow.Products[productNumber].Title);
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + (productNumber * 2 + 1)).StringCellValue,
						currentTestDataRow.Products[productNumber].Amount.ToString(CultureInfo.InvariantCulture));
				}

				for (int enumProp1Number = 0; enumProp1Number < currentTestDataRow.EnumProps1.Count; enumProp1Number++) {
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + (enumProp1Number * 3)).StringCellValue,
						currentTestDataRow.EnumProps1[enumProp1Number].Field1);
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + (enumProp1Number * 3 + 1)).StringCellValue,
						currentTestDataRow.EnumProps1[enumProp1Number].Field2.ToString(CultureInfo.InvariantCulture));
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + (enumProp1Number * 3 + 2)).StringCellValue,
						currentTestDataRow.EnumProps1[enumProp1Number].Field3.ToRussianString());
				}

				for (int enumProp2Number = 0; enumProp2Number < currentTestDataRow.EnumProps2.Count; enumProp2Number++) {
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + testDataEnumProp1ColumnsQuantity + (enumProp2Number * 5)).StringCellValue,
						currentTestDataRow.EnumProps2[enumProp2Number].Field4);
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + testDataEnumProp1ColumnsQuantity + (enumProp2Number * 5 + 1)).StringCellValue,
						currentTestDataRow.EnumProps2[enumProp2Number].Field5.ToString(CultureInfo.InvariantCulture));
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + testDataEnumProp1ColumnsQuantity + (enumProp2Number * 5 + 2)).StringCellValue,
						currentTestDataRow.EnumProps2[enumProp2Number].Field6.ToRussianString());
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + testDataEnumProp1ColumnsQuantity + (enumProp2Number * 5 + 3)).StringCellValue,
						currentTestDataRow.EnumProps2[enumProp2Number].Field7);
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + testDataEnumProp1ColumnsQuantity + (enumProp2Number * 5 + 4)).StringCellValue,
						currentTestDataRow.EnumProps2[enumProp2Number].Field8.ToString(CultureInfo.InvariantCulture));
				}
			}
		}
		[Test]
		public void ExcelStyleSimpleExport() {
			NotRelectionTestDataEntities.TestData simpleTestData = NotRelectionTestDataEntities.CreateSimpleTestData(true);
			NotRelectionTestDataEntities.ClientExampleModel firstTestDataRow = simpleTestData.Models.FirstOrDefault();
			Assert.NotNull(firstTestDataRow);
			TableWriterStyle style = new TableWriterStyle();
			MemoryStream memoryStream = TableWriterSimple.Write(new XlsTableWriterSimple(style), simpleTestData.Models, simpleTestData.ConfigurationBuilder.Value);
			WriteToFile(memoryStream, "TestStyleSimple.xls");
			HSSFWorkbook xlsFile;
			using (FileStream file = new FileStream("TestStyleSimple.xls", FileMode.Open, FileAccess.Read)) {
				xlsFile = new HSSFWorkbook(file);
			}
			ISheet sheet = xlsFile.GetSheetAt(0);

			short[] red = { 255, 0, 0 };
			short[] green = { 0, 255, 0 };
			short[] blue = { 0, 0, 255 };

			int rowNumber = 0;
			IRow row = sheet.GetRow(rowNumber);
			Assert.AreEqual(row.Height, 400);
			for (int cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
				CustomAssert.IsEqualFont(xlsFile, sheet, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Bold);
			}

			for (rowNumber = 1; rowNumber < sheet.LastRowNum; rowNumber++) {
				row = sheet.GetRow(rowNumber);
				for (int cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
					CustomAssert.IsEqualFont(xlsFile, sheet, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Normal);
				}
			}
			CustomAssert.IsEqualExcelColor((HSSFColor)sheet.GetRow(1).GetCell(2).CellStyle.FillForegroundColorColor, green);
			CustomAssert.IsEqualExcelColor((HSSFColor)sheet.GetRow(1).GetCell(15).CellStyle.FillForegroundColorColor, blue);
			CustomAssert.IsEqualExcelColor((HSSFColor)sheet.GetRow(2).GetCell(0).CellStyle.FillForegroundColorColor, red);
			CustomAssert.IsEqualExcelColor((HSSFColor)sheet.GetRow(2).GetCell(15).CellStyle.FillForegroundColorColor, blue);
		}
		[Test]
		public void ExcelComplexExport() {
			NotRelectionTestDataEntities.TestData simpleTestData = NotRelectionTestDataEntities.CreateSimpleTestData();
			NotRelectionTestDataEntities.ClientExampleModel firstTestDataRow = simpleTestData.Models.FirstOrDefault();
			Assert.NotNull(firstTestDataRow);
			TableWriterStyle style = new TableWriterStyle();
			MemoryStream memoryStream = TableWriterComplex.Write(new XlsTableWriterComplex(style), simpleTestData.Models, simpleTestData.ConfigurationBuilder.Value);
			WriteToFile(memoryStream, "TestComplex.xls");

			HSSFWorkbook xlsFile;
			using (FileStream file = new FileStream("TestComplex.xls", FileMode.Open, FileAccess.Read)) {
				xlsFile = new HSSFWorkbook(file);
			}
			ISheet sheet = xlsFile.GetSheetAt(0);
			List<string> parentColumnsNames = simpleTestData.ConfigurationBuilder.Value.ColumnsMap.Keys.ToList();

			int modelsQuantity = simpleTestData.Models.Count;
			int rowNumber = 0;
			for (int modelNumber = 0; modelNumber < modelsQuantity; modelNumber++) {
				NotRelectionTestDataEntities.ClientExampleModel currentTestDataRow = simpleTestData.Models[modelNumber];
				IRow row = sheet.GetRow(rowNumber);
				ICell cell = row.GetCell(0);
				Assert.AreEqual(cell.StringCellValue, simpleTestData.ConfigurationBuilder.Value.Title);
				for (int cellNumber = 1; cellNumber < row.LastCellNum; cellNumber++) {
					cell = row.GetCell(cellNumber);
					Assert.AreEqual(cell.StringCellValue, parentColumnsNames[cellNumber - 1]);
				}
				rowNumber++;
				row = sheet.GetRow(rowNumber);

				Assert.AreEqual(row.GetCell(1).StringCellValue, currentTestDataRow.Title);
				Assert.AreEqual(row.GetCell(2).StringCellValue, currentTestDataRow.RegistrationDate.ToRussianFullString());
				Assert.AreEqual(row.GetCell(3).StringCellValue, currentTestDataRow.Phone);
				Assert.AreEqual(row.GetCell(4).StringCellValue, currentTestDataRow.Inn);
				Assert.AreEqual(row.GetCell(5).StringCellValue, currentTestDataRow.Okato);
				Assert.AreEqual(row.GetCell(6).StringCellValue, string.Format(new CultureInfo("ru-RU"), "{0:C}", currentTestDataRow.Revenue));
				Assert.AreEqual(row.GetCell(7).StringCellValue, currentTestDataRow.EmployeeCount.ToString(CultureInfo.InvariantCulture));
				Assert.AreEqual(row.GetCell(8).StringCellValue, currentTestDataRow.IsActive.ToRussianString());
				Assert.AreEqual(row.GetCell(9).StringCellValue, currentTestDataRow.Prop1.ToString(CultureInfo.InvariantCulture));
				Assert.AreEqual(row.GetCell(10).StringCellValue, currentTestDataRow.Prop2);

				Assert.AreEqual(row.GetCell(11).StringCellValue, currentTestDataRow.Prop3.ToRussianString());
				Assert.AreEqual(row.GetCell(12).StringCellValue, currentTestDataRow.Prop4.ToString(CultureInfo.InvariantCulture));
				Assert.AreEqual(row.GetCell(13).StringCellValue, currentTestDataRow.Prop5);
				Assert.AreEqual(row.GetCell(14).StringCellValue, currentTestDataRow.Prop6.ToRussianString());
				Assert.AreEqual(row.GetCell(15).StringCellValue, currentTestDataRow.Prop7.ToString(CultureInfo.InvariantCulture));
				rowNumber++;
				row = sheet.GetRow(rowNumber);

				int childsQuantity = simpleTestData.ConfigurationBuilder.Value.ChildrenMap.Count;

				for (int childNumber = 0; childNumber < childsQuantity; childNumber++) {
					switch (childNumber) {
						case 0: {
								List<NotRelectionTestDataEntities.ClientExampleModel.Contact> child = currentTestDataRow.Contacts;
								TestChildHeader(row, childNumber, simpleTestData);
								rowNumber++;
								row = sheet.GetRow(rowNumber);
								foreach (NotRelectionTestDataEntities.ClientExampleModel.Contact childProperty in child) {
									Assert.AreEqual(row.GetCell(1).StringCellValue, childProperty.Title);
									Assert.AreEqual(row.GetCell(2).StringCellValue, childProperty.Email);
									rowNumber++;
									row = sheet.GetRow(rowNumber);
								}
								break;
							}
						case 1: {
								List<NotRelectionTestDataEntities.ClientExampleModel.Contract> child = currentTestDataRow.Contracts;
								TestChildHeader(row, childNumber, simpleTestData);
								rowNumber++;
								row = sheet.GetRow(rowNumber);
								foreach (NotRelectionTestDataEntities.ClientExampleModel.Contract childProperty in child) {
									Assert.AreEqual(row.GetCell(1).StringCellValue, childProperty.BeginDate.ToRussianFullString());
									Assert.AreEqual(row.GetCell(2).StringCellValue, childProperty.EndDate.ToRussianFullString());
									Assert.AreEqual(row.GetCell(3).StringCellValue, childProperty.Status.ToRussianString());
									rowNumber++;
									row = sheet.GetRow(rowNumber);
								}
								break;
							}
						case 2: {
								List<NotRelectionTestDataEntities.ClientExampleModel.Product> child = currentTestDataRow.Products;
								TestChildHeader(row, childNumber, simpleTestData);
								rowNumber++;
								row = sheet.GetRow(rowNumber);
								foreach (NotRelectionTestDataEntities.ClientExampleModel.Product childProperty in child) {
									Assert.AreEqual(row.GetCell(1).StringCellValue, childProperty.Title);
									Assert.AreEqual(row.GetCell(2).StringCellValue, childProperty.Amount.ToString(CultureInfo.InvariantCulture));
									rowNumber++;
									row = sheet.GetRow(rowNumber);
								}
								break;
							}
						case 3: {
								List<NotRelectionTestDataEntities.ClientExampleModel.EnumProp1> child = currentTestDataRow.EnumProps1;
								TestChildHeader(row, childNumber, simpleTestData);
								rowNumber++;
								row = sheet.GetRow(rowNumber);
								foreach (NotRelectionTestDataEntities.ClientExampleModel.EnumProp1 childProperty in child) {
									Assert.AreEqual(row.GetCell(1).StringCellValue, childProperty.Field1);
									Assert.AreEqual(row.GetCell(2).StringCellValue, childProperty.Field2.ToString(CultureInfo.InvariantCulture));
									Assert.AreEqual(row.GetCell(3).StringCellValue, childProperty.Field3.ToRussianString());
									rowNumber++;
									row = sheet.GetRow(rowNumber);
								}
								break;
							}
						case 4: {
								List<NotRelectionTestDataEntities.ClientExampleModel.EnumProp2> child = currentTestDataRow.EnumProps2;
								TestChildHeader(row, childNumber, simpleTestData);
								rowNumber++;
								row = sheet.GetRow(rowNumber);
								foreach (NotRelectionTestDataEntities.ClientExampleModel.EnumProp2 childProperty in child) {
									Assert.AreEqual(row.GetCell(1).StringCellValue, childProperty.Field4);
									Assert.AreEqual(row.GetCell(2).StringCellValue, childProperty.Field5.ToString(CultureInfo.InvariantCulture));
									Assert.AreEqual(row.GetCell(3).StringCellValue, childProperty.Field6.ToRussianString());
									Assert.AreEqual(row.GetCell(4).StringCellValue, childProperty.Field7);
									Assert.AreEqual(row.GetCell(5).StringCellValue, childProperty.Field8.ToString(CultureInfo.InvariantCulture));
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
			NotRelectionTestDataEntities.TestData simpleTestData = NotRelectionTestDataEntities.CreateSimpleTestData(true);
			NotRelectionTestDataEntities.ClientExampleModel firstTestDataRow = simpleTestData.Models.FirstOrDefault();
			Assert.NotNull(firstTestDataRow);
			TableWriterStyle style = new TableWriterStyle();
			MemoryStream memoryStream = TableWriterComplex.Write(new XlsTableWriterComplex(style), simpleTestData.Models, simpleTestData.ConfigurationBuilder.Value);
			WriteToFile(memoryStream, "TestStyleComplex.xls");

			HSSFWorkbook xlsFile;
			using (FileStream file = new FileStream("TestStyleComplex.xls", FileMode.Open, FileAccess.Read)) {
				xlsFile = new HSSFWorkbook(file);
			}
			ISheet sheet = xlsFile.GetSheetAt(0);

			short[] red = { 255, 0, 0 };
			short[] green = { 0, 255, 0 };
			short[] blue = { 0, 0, 255 };

			CustomAssert.IsEqualExcelColor((HSSFColor)sheet.GetRow(1).GetCell(3).CellStyle.FillForegroundColorColor, green);
			CustomAssert.IsEqualExcelColor((HSSFColor)sheet.GetRow(3).GetCell(1).CellStyle.FillForegroundColorColor, blue);
			CustomAssert.IsEqualExcelColor((HSSFColor)sheet.GetRow(22).GetCell(1).CellStyle.FillForegroundColorColor, red);
			CustomAssert.IsEqualExcelColor((HSSFColor)sheet.GetRow(24).GetCell(1).CellStyle.FillForegroundColorColor, blue);

			int modelsQuantity = simpleTestData.Models.Count;
			int rowNumber = 0;
			int childsQuantity = 0;
			for (int modelNumber = 0; modelNumber < modelsQuantity; modelNumber++) {
				childsQuantity += 7 + simpleTestData.Models[modelNumber].Contacts.Count + simpleTestData.Models[modelNumber].Contracts.Count + simpleTestData.Models[modelNumber].Products.Count + simpleTestData.Models[modelNumber].EnumProps1.Count + simpleTestData.Models[modelNumber].EnumProps2.Count;
				IRow row = sheet.GetRow(rowNumber);
				Assert.AreEqual(row.Height, 400);
				for (int cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
					CustomAssert.IsEqualFont(xlsFile, sheet, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Bold);
				}
				rowNumber++;

				for (; rowNumber < childsQuantity; rowNumber++) {
					row = sheet.GetRow(rowNumber);
					for (int cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
						CustomAssert.IsEqualFont(xlsFile, sheet, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Normal);
					}
				}
			}
		}
		[Test]
		public void ExcelReflectionSimpleExport() {
			List<TestDataEntities.ClientExampleModel> models = TestDataEntities.CreateSimpleTestDataModels();
			MemoryStream memoryStream = ReflectionWriterSimple.Write(models, new XlsTableWriterSimple(), new CultureInfo("ru-Ru"));
			const string fileName = "TestReflectionSimple.xls";
			WriteToFile(memoryStream, fileName);

			ExcelReflectionSimpleExportTest(models, fileName);

			List<ReflectionTestDataEntities> test = new List<ReflectionTestDataEntities>();
			Random rand = new Random();
			for (int i = 2; i < rand.Next(10, 30); i++) {
				test.Add(new ReflectionTestDataEntities());
			}
			memoryStream = ReflectionWriterSimple.Write(test, new XlsTableWriterSimple(), new CultureInfo("ru-Ru"));
			WriteToFile(memoryStream, "RandomTestReflectionSimple.xls");
			ExcelReflectionSimpleExportTest(test, "RandomTestReflectionSimple.xls");
		}
		[Test]
		public void ExcelStyleReflectionSimpleExport() {
			List<TestDataEntities.ClientExampleModel> models = TestDataEntities.CreateSimpleTestDataModels();
			TableWriterStyle style = new TableWriterStyle();
			MemoryStream memoryStream = ReflectionWriterSimple.Write(models, new XlsTableWriterSimple(style), new CultureInfo("ru-Ru"));
			const string fileName = "TestReflectionStyleSimple.xls";
			WriteToFile(memoryStream, fileName);

			ExcelStyleReflectionSimpleExportTest(models, fileName);

			List<ReflectionTestDataEntities> test = new List<ReflectionTestDataEntities>();
			Random rand = new Random();
			for (int i = 2; i < rand.Next(10, 30); i++) {
				test.Add(new ReflectionTestDataEntities());
			}
			memoryStream = ReflectionWriterSimple.Write(test, new XlsTableWriterSimple(), new CultureInfo("ru-Ru"));
			WriteToFile(memoryStream, "RandomStyleTestReflectionSimple.xls");
			ExcelStyleReflectionSimpleExportTest(test, "RandomStyleTestReflectionSimple.xls");
		}
		[Test]
		public void ExcelReflectionComplexExport() {
			List<TestDataEntities.ClientExampleModel> models = TestDataEntities.CreateSimpleTestDataModels();
			TableWriterStyle style = new TableWriterStyle();
			MemoryStream memoryStream = ReflectionWriterComplex.Write(models, new XlsTableWriterComplex(style), new CultureInfo("ru-Ru"));
			const string fileName = "TestReflectionComplex.xls";
			WriteToFile(memoryStream, fileName);

			ExcelReflectionComplexExportTest(models, fileName);

			List<ReflectionTestDataEntities> test = new List<ReflectionTestDataEntities>();
			Random rand = new Random();
			for (int i = 2; i < rand.Next(10, 30); i++) {
				test.Add(new ReflectionTestDataEntities());
			}
			memoryStream = ReflectionWriterComplex.Write(test, new XlsTableWriterComplex(style), new CultureInfo("ru-Ru"));
			WriteToFile(memoryStream, "RandomTestReflectionComplex.xls");
			ExcelReflectionComplexExportTest(test, "RandomTestReflectionComplex.xls");
		}
		[Test]
		public void ExcelStyleReflectionComplexExport() {
			List<TestDataEntities.ClientExampleModel> models = TestDataEntities.CreateSimpleTestDataModels();
			TableWriterStyle style = new TableWriterStyle();
			MemoryStream memoryStream = ReflectionWriterComplex.Write(models, new XlsTableWriterComplex(style), new CultureInfo("ru-Ru"));
			const string fileName = "TestReflectionStyleComplex.xls";
			WriteToFile(memoryStream, fileName);

			ExcelStyleReflectionComplexExportTest(models, fileName);

			List<ReflectionTestDataEntities> test = new List<ReflectionTestDataEntities>();
			Random rand = new Random();
			for (int i = 2; i < rand.Next(10, 30); i++) {
				test.Add(new ReflectionTestDataEntities());
			}
			memoryStream = ReflectionWriterComplex.Write(test, new XlsTableWriterComplex(style), new CultureInfo("ru-Ru"));
			WriteToFile(memoryStream, "RandomStyleTestReflectionComplex.xls");
			ExcelStyleReflectionComplexExportTest(test, "RandomStyleTestReflectionComplex.xls");
		}
		[Test]
		public void ExcelSimpleXlsxExport() {
			NotRelectionTestDataEntities.TestData simpleTestData = NotRelectionTestDataEntities.CreateSimpleTestData(true);
			NotRelectionTestDataEntities.ClientExampleModel firstTestDataRow = simpleTestData.Models.FirstOrDefault();
			Assert.NotNull(firstTestDataRow);
			TableWriterStyle style = new TableWriterStyle();
			TableWriterSimple.Write(new XlsxTableWriterSimple(style), simpleTestData.Models, simpleTestData.ConfigurationBuilder.Value);
			NotRelectionTestDataEntities.ClientExampleModel.Contact firstContact = firstTestDataRow.Contacts.FirstOrDefault();
			Assert.NotNull(firstContact);

			NotRelectionTestDataEntities.ClientExampleModel.Contract firstContract = firstTestDataRow.Contracts.FirstOrDefault();
			Assert.NotNull(firstContract);

			NotRelectionTestDataEntities.ClientExampleModel.Product firstProduct = firstTestDataRow.Products.FirstOrDefault();
			Assert.NotNull(firstProduct);

			NotRelectionTestDataEntities.ClientExampleModel.EnumProp1 firstEnumProp1 = firstTestDataRow.EnumProps1.FirstOrDefault();
			Assert.NotNull(firstEnumProp1);

			NotRelectionTestDataEntities.ClientExampleModel.EnumProp2 firstEnumProp2 = firstTestDataRow.EnumProps2.FirstOrDefault();
			Assert.NotNull(firstEnumProp2);

			int parentColumnsQuantity = simpleTestData.ConfigurationBuilder.Value.ColumnsMap.Count;
			int contactsFieldsQuantity = firstContact.GetType().GetProperties().Count();
			int contractsFieldsQuantity = firstContract.GetType().GetProperties().Count();
			int productsFieldsQuantity = firstProduct.GetType().GetProperties().Count();
			int enumProp1FieldsQuantity = firstEnumProp1.GetType().GetProperties().Count();

			int testDataContactColumnsQuantity = simpleTestData.Models.Max(x => x.Contacts.Count) * contactsFieldsQuantity;
			int testDataContractColumnsQuantity = simpleTestData.Models.Max(x => x.Contracts.Count) * contractsFieldsQuantity;
			int testDataProductColumnsQuantity = simpleTestData.Models.Max(x => x.Products.Count) * productsFieldsQuantity;
			int testDataEnumProp1ColumnsQuantity = simpleTestData.Models.Max(x => x.EnumProps1.Count) * enumProp1FieldsQuantity;

			List<string> parentColumnsNames = simpleTestData.ConfigurationBuilder.Value.ColumnsMap.Keys.ToList();
			List<List<string>> childsColumnsNames = simpleTestData.ConfigurationBuilder.Value.ChildrenMap.Select(childs => childs.ColumnsMap.Keys.ToList()).ToList();
			XSSFWorkbook xlsFile;
			using (FileStream file = new FileStream("TestSimple.xlsx", FileMode.Open, FileAccess.Read)) {
				xlsFile = new XSSFWorkbook(file);
			}
			ISheet sheet = xlsFile.GetSheetAt(0);
			int rowNumber = 0;
			IRow row = sheet.GetRow(rowNumber);

			int columnNumber;
			for (columnNumber = 0; columnNumber < parentColumnsNames.Count; columnNumber++) {
				Assert.AreEqual(row.GetCell(columnNumber).StringCellValue, parentColumnsNames[columnNumber]);
			}

			columnNumber = parentColumnsQuantity;
			for (int childNumber = 0; childNumber < childsColumnsNames.Count; childNumber++) {
				List<string> child = childsColumnsNames[childNumber];
				int childColumnsQuantity = 0;
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
				for (int index = 1; index <= childColumnsQuantity; index++) {
					foreach (string childPropertyName in child) {
						Assert.AreEqual(row.GetCell(columnNumber).StringCellValue, childPropertyName + index);
						columnNumber++;
					}
				}
			}
			for (rowNumber = 1; rowNumber <= sheet.LastRowNum; rowNumber++) {
				row = sheet.GetRow(rowNumber);
				Assert.NotNull(row);
				NotRelectionTestDataEntities.ClientExampleModel currentTestDataRow = simpleTestData.Models[rowNumber - 1];
				Assert.AreEqual(row.GetCell(0).StringCellValue, currentTestDataRow.Title);
				Assert.AreEqual(row.GetCell(1).StringCellValue, currentTestDataRow.RegistrationDate.ToRussianFullString());
				Assert.AreEqual(row.GetCell(2).StringCellValue, currentTestDataRow.Phone);
				Assert.AreEqual(row.GetCell(3).StringCellValue, currentTestDataRow.Inn);
				Assert.AreEqual(row.GetCell(4).StringCellValue, currentTestDataRow.Okato);

				Assert.AreEqual(row.GetCell(5).StringCellValue, string.Format(new CultureInfo("ru-RU"), "{0:C}", currentTestDataRow.Revenue));
				Assert.AreEqual(row.GetCell(6).StringCellValue, currentTestDataRow.EmployeeCount.ToString(CultureInfo.InvariantCulture));
				Assert.AreEqual(row.GetCell(7).StringCellValue, currentTestDataRow.IsActive.ToRussianString());
				Assert.AreEqual(row.GetCell(8).StringCellValue, currentTestDataRow.Prop1.ToString(CultureInfo.InvariantCulture));
				Assert.AreEqual(row.GetCell(9).StringCellValue, currentTestDataRow.Prop2);

				Assert.AreEqual(row.GetCell(10).StringCellValue, currentTestDataRow.Prop3.ToRussianString());
				Assert.AreEqual(row.GetCell(11).StringCellValue, currentTestDataRow.Prop4.ToString(CultureInfo.InvariantCulture));
				Assert.AreEqual(row.GetCell(12).StringCellValue, currentTestDataRow.Prop5);
				Assert.AreEqual(row.GetCell(13).StringCellValue, currentTestDataRow.Prop6.ToRussianString());
				Assert.AreEqual(row.GetCell(14).StringCellValue, currentTestDataRow.Prop7.ToString(CultureInfo.InvariantCulture));

				for (int contactNumber = 0; contactNumber < currentTestDataRow.Contacts.Count; contactNumber++) {
					NotRelectionTestDataEntities.ClientExampleModel.Contact currentContactRow = currentTestDataRow.Contacts[contactNumber];
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + (contactNumber * contactsFieldsQuantity)).StringCellValue, currentContactRow.Title);
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + (contactNumber * contactsFieldsQuantity + 1)).StringCellValue, currentContactRow.Email);
				}

				for (int contractNumber = 0; contractNumber < currentTestDataRow.Contracts.Count; contractNumber++) {
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + (contractNumber * contractsFieldsQuantity)).StringCellValue,
						currentTestDataRow.Contracts[contractNumber].BeginDate.ToRussianFullString());
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + (contractNumber * contractsFieldsQuantity + 1)).StringCellValue,
						currentTestDataRow.Contracts[contractNumber].EndDate.ToRussianFullString());
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + (contractNumber * contractsFieldsQuantity + 2)).StringCellValue,
						currentTestDataRow.Contracts[contractNumber].Status.ToRussianString());
				}

				for (int productNumber = 0; productNumber < currentTestDataRow.Products.Count; productNumber++) {
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + (productNumber * 2)).StringCellValue,
						currentTestDataRow.Products[productNumber].Title);
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + (productNumber * 2 + 1)).StringCellValue,
						currentTestDataRow.Products[productNumber].Amount.ToString(CultureInfo.InvariantCulture));
				}

				for (int enumProp1Number = 0; enumProp1Number < currentTestDataRow.EnumProps1.Count; enumProp1Number++) {
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + (enumProp1Number * 3)).StringCellValue,
						currentTestDataRow.EnumProps1[enumProp1Number].Field1);
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + (enumProp1Number * 3 + 1)).StringCellValue,
						currentTestDataRow.EnumProps1[enumProp1Number].Field2.ToString(CultureInfo.InvariantCulture));
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + (enumProp1Number * 3 + 2)).StringCellValue,
						currentTestDataRow.EnumProps1[enumProp1Number].Field3.ToRussianString());
				}

				for (int enumProp2Number = 0; enumProp2Number < currentTestDataRow.EnumProps2.Count; enumProp2Number++) {
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + testDataEnumProp1ColumnsQuantity + (enumProp2Number * 5)).StringCellValue,
						currentTestDataRow.EnumProps2[enumProp2Number].Field4);
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + testDataEnumProp1ColumnsQuantity + (enumProp2Number * 5 + 1)).StringCellValue,
						currentTestDataRow.EnumProps2[enumProp2Number].Field5.ToString(CultureInfo.InvariantCulture));
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + testDataEnumProp1ColumnsQuantity + (enumProp2Number * 5 + 2)).StringCellValue,
						currentTestDataRow.EnumProps2[enumProp2Number].Field6.ToRussianString());
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + testDataEnumProp1ColumnsQuantity + (enumProp2Number * 5 + 3)).StringCellValue,
						currentTestDataRow.EnumProps2[enumProp2Number].Field7);
					Assert.AreEqual(row.GetCell(parentColumnsQuantity + testDataContactColumnsQuantity + testDataContractColumnsQuantity + testDataProductColumnsQuantity + testDataEnumProp1ColumnsQuantity + (enumProp2Number * 5 + 4)).StringCellValue,
						currentTestDataRow.EnumProps2[enumProp2Number].Field8.ToString(CultureInfo.InvariantCulture));
				}
			}
		}
		[Test]
		public void ExcelStyleSimpleXlsxExport() {
			NotRelectionTestDataEntities.TestData simpleTestData = NotRelectionTestDataEntities.CreateSimpleTestData(true);
			NotRelectionTestDataEntities.ClientExampleModel firstTestDataRow = simpleTestData.Models.FirstOrDefault();
			Assert.NotNull(firstTestDataRow);
			TableWriterStyle style = new TableWriterStyle();
			TableWriterSimple.Write(new XlsxTableWriterSimple(style), simpleTestData.Models, simpleTestData.ConfigurationBuilder.Value);
			NotRelectionTestDataEntities.ClientExampleModel.Contact firstContact = firstTestDataRow.Contacts.FirstOrDefault();
			Assert.NotNull(firstContact);

			NotRelectionTestDataEntities.ClientExampleModel.Contract firstContract = firstTestDataRow.Contracts.FirstOrDefault();
			Assert.NotNull(firstContract);

			NotRelectionTestDataEntities.ClientExampleModel.Product firstProduct = firstTestDataRow.Products.FirstOrDefault();
			Assert.NotNull(firstProduct);

			NotRelectionTestDataEntities.ClientExampleModel.EnumProp1 firstEnumProp1 = firstTestDataRow.EnumProps1.FirstOrDefault();
			Assert.NotNull(firstEnumProp1);

			NotRelectionTestDataEntities.ClientExampleModel.EnumProp2 firstEnumProp2 = firstTestDataRow.EnumProps2.FirstOrDefault();
			Assert.NotNull(firstEnumProp2);

			XSSFWorkbook xlsFile;
			using (FileStream file = new FileStream("TestSimple.xlsx", FileMode.Open, FileAccess.Read)) {
				xlsFile = new XSSFWorkbook(file);
			}
			ISheet sheet = xlsFile.GetSheetAt(0);

			short[] red = { 255, 0, 0 };
			short[] green = { 0, 255, 0 };
			short[] blue = { 0, 0, 255 };

			int rowNumber = 0;
			IRow row = sheet.GetRow(rowNumber);
			Assert.AreEqual(row.Height, 400);
			for (int cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
				CustomAssert.IsEqualFont(xlsFile, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Bold);
			}

			for (rowNumber = 1; rowNumber < sheet.LastRowNum; rowNumber++) {
				row = sheet.GetRow(rowNumber);
				for (int cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
					CustomAssert.IsEqualFont(xlsFile, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Normal);
				}
			}
			CustomAssert.IsEqualExcelColor((XSSFColor)sheet.GetRow(1).GetCell(2).CellStyle.FillForegroundColorColor, green);
			CustomAssert.IsEqualExcelColor((XSSFColor)sheet.GetRow(1).GetCell(15).CellStyle.FillForegroundColorColor, blue);
			CustomAssert.IsEqualExcelColor((XSSFColor)sheet.GetRow(2).GetCell(0).CellStyle.FillForegroundColorColor, red);
			CustomAssert.IsEqualExcelColor((XSSFColor)sheet.GetRow(2).GetCell(15).CellStyle.FillForegroundColorColor, blue);
		}
		[Test]
		public void ExcelComplexXlsxExport() {
			NotRelectionTestDataEntities.TestData simpleTestData = NotRelectionTestDataEntities.CreateSimpleTestData(true);
			NotRelectionTestDataEntities.ClientExampleModel firstTestDataRow = simpleTestData.Models.FirstOrDefault();
			Assert.NotNull(firstTestDataRow);
			TableWriterStyle style = new TableWriterStyle();
			TableWriterComplex.Write(new XlsxTableWriterComplex(style), simpleTestData.Models, simpleTestData.ConfigurationBuilder.Value);

			XSSFWorkbook xlsFile;
			using (FileStream file = new FileStream("TestComplex.xlsx", FileMode.Open, FileAccess.Read)) {
				xlsFile = new XSSFWorkbook(file);
			}
			ISheet sheet = xlsFile.GetSheetAt(0);
			List<string> parentColumnsNames = simpleTestData.ConfigurationBuilder.Value.ColumnsMap.Keys.ToList();

			int modelsQuantity = simpleTestData.Models.Count;
			int rowNumber = 0;
			for (int modelNumber = 0; modelNumber < modelsQuantity; modelNumber++) {
				NotRelectionTestDataEntities.ClientExampleModel currentTestDataRow = simpleTestData.Models[modelNumber];
				IRow row = sheet.GetRow(rowNumber);
				ICell cell = row.GetCell(0);
				Assert.AreEqual(cell.StringCellValue, simpleTestData.ConfigurationBuilder.Value.Title);
				for (int cellNumber = 1; cellNumber < row.LastCellNum; cellNumber++) {
					cell = row.GetCell(cellNumber);
					Assert.AreEqual(cell.StringCellValue, parentColumnsNames[cellNumber - 1]);
				}
				rowNumber++;
				row = sheet.GetRow(rowNumber);

				Assert.AreEqual(row.GetCell(1).StringCellValue, currentTestDataRow.Title);
				Assert.AreEqual(row.GetCell(2).StringCellValue, currentTestDataRow.RegistrationDate.ToRussianFullString());
				Assert.AreEqual(row.GetCell(3).StringCellValue, currentTestDataRow.Phone);
				Assert.AreEqual(row.GetCell(4).StringCellValue, currentTestDataRow.Inn);
				Assert.AreEqual(row.GetCell(5).StringCellValue, currentTestDataRow.Okato);
				Assert.AreEqual(row.GetCell(6).StringCellValue, string.Format(new CultureInfo("ru-RU"), "{0:C}", currentTestDataRow.Revenue));
				Assert.AreEqual(row.GetCell(7).StringCellValue, currentTestDataRow.EmployeeCount.ToString(CultureInfo.InvariantCulture));
				Assert.AreEqual(row.GetCell(8).StringCellValue, currentTestDataRow.IsActive.ToRussianString());
				Assert.AreEqual(row.GetCell(9).StringCellValue, currentTestDataRow.Prop1.ToString(CultureInfo.InvariantCulture));
				Assert.AreEqual(row.GetCell(10).StringCellValue, currentTestDataRow.Prop2);

				Assert.AreEqual(row.GetCell(11).StringCellValue, currentTestDataRow.Prop3.ToRussianString());
				Assert.AreEqual(row.GetCell(12).StringCellValue, currentTestDataRow.Prop4.ToString(CultureInfo.InvariantCulture));
				Assert.AreEqual(row.GetCell(13).StringCellValue, currentTestDataRow.Prop5);
				Assert.AreEqual(row.GetCell(14).StringCellValue, currentTestDataRow.Prop6.ToRussianString());
				Assert.AreEqual(row.GetCell(15).StringCellValue, currentTestDataRow.Prop7.ToString(CultureInfo.InvariantCulture));
				rowNumber++;
				row = sheet.GetRow(rowNumber);

				int childsQuantity = simpleTestData.ConfigurationBuilder.Value.ChildrenMap.Count;

				for (int childNumber = 0; childNumber < childsQuantity; childNumber++) {
					switch (childNumber) {
						case 0: {
								List<NotRelectionTestDataEntities.ClientExampleModel.Contact> child = currentTestDataRow.Contacts;
								TestChildHeader(row, childNumber, simpleTestData);
								rowNumber++;
								row = sheet.GetRow(rowNumber);
								foreach (NotRelectionTestDataEntities.ClientExampleModel.Contact childProperty in child) {
									Assert.AreEqual(row.GetCell(1).StringCellValue, childProperty.Title);
									Assert.AreEqual(row.GetCell(2).StringCellValue, childProperty.Email);
									rowNumber++;
									row = sheet.GetRow(rowNumber);
								}
								break;
							}
						case 1: {
								List<NotRelectionTestDataEntities.ClientExampleModel.Contract> child = currentTestDataRow.Contracts;
								TestChildHeader(row, childNumber, simpleTestData);
								rowNumber++;
								row = sheet.GetRow(rowNumber);
								foreach (NotRelectionTestDataEntities.ClientExampleModel.Contract childProperty in child) {
									Assert.AreEqual(row.GetCell(1).StringCellValue, childProperty.BeginDate.ToRussianFullString());
									Assert.AreEqual(row.GetCell(2).StringCellValue, childProperty.EndDate.ToRussianFullString());
									Assert.AreEqual(row.GetCell(3).StringCellValue, childProperty.Status.ToRussianString());
									rowNumber++;
									row = sheet.GetRow(rowNumber);
								}
								break;
							}
						case 2: {
								List<NotRelectionTestDataEntities.ClientExampleModel.Product> child = currentTestDataRow.Products;
								TestChildHeader(row, childNumber, simpleTestData);
								rowNumber++;
								row = sheet.GetRow(rowNumber);
								foreach (NotRelectionTestDataEntities.ClientExampleModel.Product childProperty in child) {
									Assert.AreEqual(row.GetCell(1).StringCellValue, childProperty.Title);
									Assert.AreEqual(row.GetCell(2).StringCellValue, childProperty.Amount.ToString(CultureInfo.InvariantCulture));
									rowNumber++;
									row = sheet.GetRow(rowNumber);
								}
								break;
							}
						case 3: {
								List<NotRelectionTestDataEntities.ClientExampleModel.EnumProp1> child = currentTestDataRow.EnumProps1;
								TestChildHeader(row, childNumber, simpleTestData);
								rowNumber++;
								row = sheet.GetRow(rowNumber);
								foreach (NotRelectionTestDataEntities.ClientExampleModel.EnumProp1 childProperty in child) {
									Assert.AreEqual(row.GetCell(1).StringCellValue, childProperty.Field1);
									Assert.AreEqual(row.GetCell(2).StringCellValue, childProperty.Field2.ToString(CultureInfo.InvariantCulture));
									Assert.AreEqual(row.GetCell(3).StringCellValue, childProperty.Field3.ToRussianString());
									rowNumber++;
									row = sheet.GetRow(rowNumber);
								}
								break;
							}
						case 4: {
								List<NotRelectionTestDataEntities.ClientExampleModel.EnumProp2> child = currentTestDataRow.EnumProps2;
								TestChildHeader(row, childNumber, simpleTestData);
								rowNumber++;
								row = sheet.GetRow(rowNumber);
								foreach (NotRelectionTestDataEntities.ClientExampleModel.EnumProp2 childProperty in child) {
									Assert.AreEqual(row.GetCell(1).StringCellValue, childProperty.Field4);
									Assert.AreEqual(row.GetCell(2).StringCellValue, childProperty.Field5.ToString(CultureInfo.InvariantCulture));
									Assert.AreEqual(row.GetCell(3).StringCellValue, childProperty.Field6.ToRussianString());
									Assert.AreEqual(row.GetCell(4).StringCellValue, childProperty.Field7);
									Assert.AreEqual(row.GetCell(5).StringCellValue, childProperty.Field8.ToString(CultureInfo.InvariantCulture));
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
			NotRelectionTestDataEntities.TestData simpleTestData = NotRelectionTestDataEntities.CreateSimpleTestData(true);
			NotRelectionTestDataEntities.ClientExampleModel firstTestDataRow = simpleTestData.Models.FirstOrDefault();
			Assert.NotNull(firstTestDataRow);
			TableWriterStyle style = new TableWriterStyle();
			TableWriterComplex.Write(new XlsxTableWriterComplex(style), simpleTestData.Models, simpleTestData.ConfigurationBuilder.Value);

			XSSFWorkbook xlsFile;
			using (FileStream file = new FileStream("TestComplex.xlsx", FileMode.Open, FileAccess.Read)) {
				xlsFile = new XSSFWorkbook(file);
			}
			ISheet sheet = xlsFile.GetSheetAt(0);

			short[] red = { 255, 0, 0 };
			short[] green = { 0, 255, 0 };
			short[] blue = { 0, 0, 255 };

			CustomAssert.IsEqualExcelColor((XSSFColor)sheet.GetRow(1).GetCell(3).CellStyle.FillForegroundColorColor, green);
			CustomAssert.IsEqualExcelColor((XSSFColor)sheet.GetRow(3).GetCell(1).CellStyle.FillForegroundColorColor, blue);
			CustomAssert.IsEqualExcelColor((XSSFColor)sheet.GetRow(22).GetCell(1).CellStyle.FillForegroundColorColor, red);
			CustomAssert.IsEqualExcelColor((XSSFColor)sheet.GetRow(24).GetCell(1).CellStyle.FillForegroundColorColor, blue);

			int modelsQuantity = simpleTestData.Models.Count;
			int rowNumber = 0;
			int childsQuantity = 0;
			for (int modelNumber = 0; modelNumber < modelsQuantity; modelNumber++) {
				childsQuantity += 7 + simpleTestData.Models[modelNumber].Contacts.Count + simpleTestData.Models[modelNumber].Contracts.Count + simpleTestData.Models[modelNumber].Products.Count + simpleTestData.Models[modelNumber].EnumProps1.Count + simpleTestData.Models[modelNumber].EnumProps2.Count;
				IRow row = sheet.GetRow(rowNumber);
				Assert.AreEqual(row.Height, 400);
				for (int cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
					CustomAssert.IsEqualFont(xlsFile, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Bold);
				}
				rowNumber++;

				for (; rowNumber < childsQuantity; rowNumber++) {
					row = sheet.GetRow(rowNumber);
					for (int cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
						CustomAssert.IsEqualFont(xlsFile, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Normal);
					}
				}
			}
		}
		[Test]
		public void ExcelReflectionSimpleXlsxExport() {
			List<ReflectionTestDataEntities> test = new List<ReflectionTestDataEntities>();
			TableWriterStyle style = new TableWriterStyle();
			Random rand = new Random();
			for (int i = 2; i < rand.Next(10, 30); i++) {
				test.Add(new ReflectionTestDataEntities());
			}
			ReflectionWriterSimple.Write(test, new XlsxTableWriterSimple(style), new CultureInfo("ru-Ru"));
			ExcelReflectionSimpleExportTest(test, "TestSimple.xlsx", true);

			ExcelStyleReflectionSimpleExportTest(test, "TestSimple.xlsx", true);
		}
		[Test]
		public void ExcelReflectionComplexXlsxExport() {
			List<ReflectionTestDataEntities> test = new List<ReflectionTestDataEntities>();
			TableWriterStyle style = new TableWriterStyle();
			Random rand = new Random();
			for (int i = 2; i < rand.Next(10, 30); i++) {
				test.Add(new ReflectionTestDataEntities());
			}
			ReflectionWriterComplex.Write(test, new XlsxTableWriterComplex(style), new CultureInfo("ru-Ru"));
			ExcelReflectionComplexExportTest(test, "TestComplex.xlsx", true);

			ExcelStyleReflectionComplexExportTest(test, "TestComplex.xlsx", true);
		}
		private static void ExcelReflectionSimpleExportTest<T>(List<T> models, string fileName, bool isXlsx = false) {
			T firstModel = models.FirstOrDefault();
			Assert.NotNull(firstModel);
			List<Type> generalTypes = GeneralTypes;
			List<PropertyInfo> nonEnumerableProperties = firstModel.GetType().GetProperties().Where(x => generalTypes.Contains(x.PropertyType)).ToList();
			List<PropertyInfo> enumerableProperties = firstModel.GetType().GetProperties().Where(x => x.PropertyType.IsGenericType && x.PropertyType.GetGenericTypeDefinition() == typeof(List<>)).ToList();

			IWorkbook xlsFile;
			using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read)) {
				if (isXlsx) {
					xlsFile = new XSSFWorkbook(file);
				} else {
					xlsFile = new HSSFWorkbook(file);
				}
			}
			ISheet sheet = xlsFile.GetSheetAt(0);
			int rowNumber = 0;
			IRow row = sheet.GetRow(rowNumber);
			int cellNumber = 0;

			int[] enumerablePropertiesChildrensMaxQuantity = new int[enumerableProperties.Count()];
			for (int i = 0; i < enumerablePropertiesChildrensMaxQuantity.Count(); i++) {
				enumerablePropertiesChildrensMaxQuantity[i] = 0;
			}
			foreach (T model in models) {
				for (int i = 0; i < enumerableProperties.Count(); i++) {
					object nestedModels = enumerableProperties.ElementAt(i).GetValue(model);
					if (enumerablePropertiesChildrensMaxQuantity[i] < ((IList)nestedModels).Count) {
						enumerablePropertiesChildrensMaxQuantity[i] = ((IList)nestedModels).Count;
					}
				}
			}
			foreach (PropertyInfo property in nonEnumerableProperties) {
				ExcelExportAttribute attribute = property.GetCustomAttribute<ExcelExportAttribute>();
				if (attribute != null && attribute.IsExportable) {
					Assert.AreEqual(row.GetCell(cellNumber).StringCellValue, attribute.PropertyName);
					cellNumber++;
				}
			}

			for (int i = 0; i < enumerableProperties.Count(); i++) {
				PropertyInfo property = enumerableProperties.ElementAt(i);
				Type propertyType = property.PropertyType;
				Type listType = propertyType.GetGenericArguments()[0];

				List<PropertyInfo> props = listType.GetProperties().Where(x => generalTypes.Contains(x.PropertyType)).ToList();
				for (int j = 0; j < enumerablePropertiesChildrensMaxQuantity[i]; j++) {
					foreach (PropertyInfo prop in props) {
						ExcelExportAttribute attribute = prop.GetCustomAttribute<ExcelExportAttribute>();
						Assert.AreEqual(row.GetCell(cellNumber).StringCellValue, attribute.PropertyName + (j + 1));
						cellNumber++;
					}
				}
			}
			CultureInfo cultureInfo = new CultureInfo("ru-RU");
			rowNumber = 1;
			foreach (T model in models) {
				cellNumber = 0;
				row = sheet.GetRow(rowNumber);
				foreach (PropertyInfo nonEnumerableProperty in nonEnumerableProperties) {
					ExcelExportAttribute attribute = nonEnumerableProperty.GetCustomAttribute<ExcelExportAttribute>();
					if (attribute != null && attribute.IsExportable) {
						string value = ConvertPropertyToString(nonEnumerableProperty, model, cultureInfo);
						Assert.AreEqual(row.GetCell(cellNumber).StringCellValue, value);
						cellNumber++;
					}
				}

				for (int i = 0; i < enumerableProperties.Count(); i++) {
					PropertyInfo property = enumerableProperties.ElementAt(i);
					IList submodels = (IList)property.GetValue(model);

					Type propertyType = property.PropertyType;
					Type listType = propertyType.GetGenericArguments()[0];

					List<PropertyInfo> listPeroperties = listType.GetProperties().Where(x => generalTypes.Contains(x.PropertyType)).ToList();

					foreach (object submodel in submodels) {
						foreach (PropertyInfo listProperty in listPeroperties) {
							ExcelExportAttribute attribute = listProperty.GetCustomAttribute<ExcelExportAttribute>();
							if (attribute != null && attribute.IsExportable) {
								string value = ConvertPropertyToString(listProperty, submodel, cultureInfo);
								Assert.AreEqual(row.GetCell(cellNumber).StringCellValue, value);
								cellNumber++;
							}
						}
					}
					cellNumber += (enumerablePropertiesChildrensMaxQuantity[i] - submodels.Count) * listPeroperties.Count();
				}
				rowNumber++;
			}
		}
		private static void ExcelStyleReflectionSimpleExportTest<T>(IEnumerable<T> models, string fileName, bool isXlsx = false) {
			T firstModel = models.FirstOrDefault();
			Assert.NotNull(firstModel);

			IWorkbook xlsFile;
			using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read)) {
				if (isXlsx) {
					xlsFile = new XSSFWorkbook(file);
				} else {
					xlsFile = new HSSFWorkbook(file);
				}
			}
			ISheet sheet = xlsFile.GetSheetAt(0);

			int rowNumber = 0;
			IRow row = sheet.GetRow(rowNumber);
			Assert.AreEqual(row.Height, 400);
			for (int cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
				if (isXlsx) {
					CustomAssert.IsEqualFont((XSSFWorkbook)xlsFile, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Bold);
				} else {
					CustomAssert.IsEqualFont(xlsFile, sheet, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Bold);
				}
			}

			for (rowNumber = 1; rowNumber < sheet.LastRowNum; rowNumber++) {
				row = sheet.GetRow(rowNumber);
				for (int cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
					if (isXlsx) {
						CustomAssert.IsEqualFont((XSSFWorkbook)xlsFile, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Normal);
					} else {
						CustomAssert.IsEqualFont(xlsFile, sheet, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Normal);
					}
				}
			}
		}
		private static void ExcelReflectionComplexExportTest<T>(List<T> models, string fileName, bool isXlsx = false) {
			T firstModel = models.FirstOrDefault();
			Assert.NotNull(firstModel);
			List<Type> generalTypes = GeneralTypes;
			List<PropertyInfo> nonEnumerableProperties = firstModel.GetType().GetProperties().Where(x => generalTypes.Contains(x.PropertyType)).ToList();
			List<PropertyInfo> enumerableProperties = firstModel.GetType().GetProperties().Where(x => x.PropertyType.IsGenericType && x.PropertyType.GetGenericTypeDefinition() == typeof(List<>)).ToList();

			IWorkbook xlsFile;
			using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read)) {
				if (isXlsx) {
					xlsFile = new XSSFWorkbook(file);
				} else {
					xlsFile = new HSSFWorkbook(file);
				}
			}
			ISheet sheet = xlsFile.GetSheetAt(0);
			int rowNumber = 0;
			IRow row = sheet.GetRow(rowNumber);
			int cellNumber = 0;
			CultureInfo cultureInfo = new CultureInfo("ru-RU");
			foreach (T model in models) {
				Assert.AreEqual(row.GetCell(cellNumber).StringCellValue, model.GetType().GetCustomAttribute<ExcelExportClassNameAttribute>().Name);
				cellNumber++;

				foreach (PropertyInfo nonEnumerableProperty in nonEnumerableProperties) {
					ExcelExportAttribute attribute = nonEnumerableProperty.GetCustomAttribute<ExcelExportAttribute>();
					if (attribute != null && attribute.IsExportable) {
						Assert.AreEqual(row.GetCell(cellNumber).StringCellValue, attribute.PropertyName);
						cellNumber++;
					}
				}
				rowNumber++;
				cellNumber = 1;

				row = sheet.GetRow(rowNumber);
				foreach (PropertyInfo nonEnumerableProperty in nonEnumerableProperties) {
					ExcelExportAttribute attribute = nonEnumerableProperty.GetCustomAttribute<ExcelExportAttribute>();
					if (attribute != null && attribute.IsExportable) {
						string value = ConvertPropertyToString(nonEnumerableProperty, model, cultureInfo);
						Assert.AreEqual(row.GetCell(cellNumber).StringCellValue, value);
						cellNumber++;
					}
				}
				rowNumber++;
				cellNumber = 0;

				row = sheet.GetRow(rowNumber);
				foreach (PropertyInfo property in enumerableProperties) {
					Type propertyType = property.PropertyType;
					Type listType = propertyType.GetGenericArguments()[0];

					List<PropertyInfo> props = listType.GetProperties().Where(x => generalTypes.Contains(x.PropertyType)).ToList();
					IList submodels = (IList)property.GetValue(model);
					if (submodels.Count != 0) {
						string submodelName = submodels[0].GetType().GetCustomAttribute<ExcelExportClassNameAttribute>().Name;
						Assert.AreEqual(row.GetCell(cellNumber).StringCellValue, submodelName);
					}
					cellNumber++;
					foreach (PropertyInfo prop in props) {
						var attribute1 = prop.GetCustomAttribute<ExcelExportAttribute>();
						Assert.AreEqual(row.GetCell(cellNumber).StringCellValue, attribute1.PropertyName);
						cellNumber++;
					}
					rowNumber++;
					row = sheet.GetRow(rowNumber);
					if (((IList)property.GetValue(model)).Count != 0) {
						foreach (object submodel in submodels) {
							cellNumber = 1;
							foreach (PropertyInfo prop in props) {
								ExcelExportAttribute attribute = prop.GetCustomAttribute<ExcelExportAttribute>();
								if (attribute != null && attribute.IsExportable) {
									string value = ConvertPropertyToString(prop, submodel, cultureInfo);
									Assert.AreEqual(row.GetCell(cellNumber).StringCellValue, value);
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
		private static void ExcelStyleReflectionComplexExportTest<T>(List<T> models, string fileName, bool isXlsx = false) {
			Assert.NotNull(models.FirstOrDefault());
			IWorkbook xlsFile;
			using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read)) {
				if (isXlsx) {
					xlsFile = new XSSFWorkbook(file);
				} else {
					xlsFile = new HSSFWorkbook(file);
				}
			}
			ISheet sheet = xlsFile.GetSheetAt(0);

			int rowNumber = 0;
			IRow row = sheet.GetRow(rowNumber);
			foreach (T model in models) {
				int cellNumber;
				Assert.AreEqual(row.Height, 400);
				for (cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
					if (isXlsx) {
						CustomAssert.IsEqualFont((XSSFWorkbook)xlsFile, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Bold);
					} else {
						CustomAssert.IsEqualFont(xlsFile, sheet, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Bold);
					}

				}
				rowNumber++;
				row = sheet.GetRow(rowNumber);
				List<PropertyInfo> enumerableProperties = model.GetType().GetProperties().Where(x => x.PropertyType.IsGenericType && x.PropertyType.GetGenericTypeDefinition() == typeof(List<>)).ToList();
				for (cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
					if (isXlsx) {
						CustomAssert.IsEqualFont((XSSFWorkbook)xlsFile, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Normal);
					} else {
						CustomAssert.IsEqualFont(xlsFile, sheet, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Normal);
					}
				}
				rowNumber++;
				row = sheet.GetRow(rowNumber);
				foreach (PropertyInfo enumerableProperty in enumerableProperties) {
					IList childs = (IList)enumerableProperty.GetValue(model);
					for (cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
						if (isXlsx) {
							CustomAssert.IsEqualFont((XSSFWorkbook)xlsFile, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Normal);
						} else {
							CustomAssert.IsEqualFont(xlsFile, sheet, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Normal);
						}
					}
					rowNumber++;

					int rowNumberStart = rowNumber;
					for (rowNumber = rowNumberStart; rowNumber < rowNumberStart + childs.Count; rowNumber++) {
						row = sheet.GetRow(rowNumber);
						for (cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
							if (isXlsx) {
								CustomAssert.IsEqualFont((XSSFWorkbook)xlsFile, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Normal);
							} else {
								CustomAssert.IsEqualFont(xlsFile, sheet, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Normal);
							}
						}
					}
					row = sheet.GetRow(rowNumber);
				}
			}
		}
		private static string ConvertPropertyToString<T>(PropertyInfo nonEnumerableProperty, T model, CultureInfo cultureInfo) {
			string propertyTypeName = nonEnumerableProperty.PropertyType.Name;
			string value = String.Empty;
			switch (propertyTypeName) {
				case "String":
					value = nonEnumerableProperty.GetValue(model).ToString();
					break;
				case "DateTime":
					value = ((DateTime)nonEnumerableProperty.GetValue(model)).ToString(cultureInfo.DateTimeFormat.LongDatePattern);
					break;
				case "Decimal":
					value = string.Format(cultureInfo, "{0:C}", nonEnumerableProperty.GetValue(model));
					break;
				case "Int32":
					value = nonEnumerableProperty.GetValue(model).ToString();
					break;
				case "Boolean":
					value = ((bool)nonEnumerableProperty.GetValue(model)) ? "Да" : "Нет";
					break;
			}
			return value;
		}
		private static List<string> TestChildHeader(IRow row, int childNumber, NotRelectionTestDataEntities.TestData simpleTestData) {
			string childName = simpleTestData.ConfigurationBuilder.Value.ChildrenMap[childNumber].Title;
			List<string> childColumnsNames = simpleTestData.ConfigurationBuilder.Value.ChildrenMap[childNumber].ColumnsMap.Keys.ToList();

			Assert.AreEqual(row.GetCell(0).StringCellValue, childName);
			for (int cellNumber = 1; cellNumber < row.LastCellNum; cellNumber++) {
				int numberProperty = cellNumber - 1;
				Assert.AreEqual(row.GetCell(cellNumber).StringCellValue, childColumnsNames[numberProperty]);
			}

			return childColumnsNames;
		}
		private static void WriteToFile(MemoryStream ms, string fileName) {
			using (FileStream file = new FileStream(fileName, FileMode.Create, FileAccess.Write)) {
				byte[] bytes = new byte[ms.Length];
				ms.Read(bytes, 0, (int)ms.Length);
				file.Write(bytes, 0, bytes.Length);
				ms.Close();
			}
		}
		private enum FontBoldWeight {
			Normal = 400,
			Bold = 700,
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