﻿using System.Collections.Generic;
using System.Linq;
using Extensions;
using FormattedExcelExport.Configuaration;
using FormattedExcelExport.Style;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NUnit.Framework;
using System.IO;
using FormattedExcelExport.TableWriters;

namespace FormattedExcelExport.Tests {
    [TestFixture]
    public class Test {
        [Test]
        public void ExcelSimpleExport() {
            TestDataEntities.TestData simpleTestData = TestDataEntities.CreateSimpleTestData();
            TestDataEntities.ClientExampleModel firstTestDataRow = simpleTestData.Models.FirstOrDefault();
            Assert.NotNull(firstTestDataRow);
            MemoryStream memoryStream = TableWriterSimple.Write(new XlsTableWriterSimple(), simpleTestData.Models, simpleTestData.ConfigurationBuilder.Value);
            WriteToFile(memoryStream, "TestSimple.xls");

            TestDataEntities.ClientExampleModel.Contact firstContact = firstTestDataRow.Contacts.FirstOrDefault();
            Assert.NotNull(firstContact);

            TestDataEntities.ClientExampleModel.Contract firstContract = firstTestDataRow.Contracts.FirstOrDefault();
            Assert.NotNull(firstContract);

            int parentColumnsQuantity = simpleTestData.ConfigurationBuilder.Value.ColumnsMap.Count;
            int contactsFieldsQuantity = firstContact.GetType().GetProperties().Count();
            int contractsFieldsQuantity = firstContract.GetType().GetProperties().Count();
            int testDataContactColumnsQuantity = simpleTestData.Models.Max(x => x.Contacts.Count) * contactsFieldsQuantity;
            int testDataContractColumnsQuantity = simpleTestData.Models.Max(x => x.Contracts.Count) * contractsFieldsQuantity;

            List<string> parentColumnsNames = simpleTestData.ConfigurationBuilder.Value.ColumnsMap.Keys.ToList();
            List<List<string>> childsColumnsNames = new List<List<string>>();
            foreach (ChildTableConfiguration childs in simpleTestData.ConfigurationBuilder.Value.ChildrenMap) {
                childsColumnsNames.Add(childs.ColumnsMap.Keys.ToList());
            }

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
                }
                for (int index = 1; index <= childColumnsQuantity; index++) {
                    for (int childPropertyNumber = 0; childPropertyNumber < child.Count; childPropertyNumber++) {
                        string childPropertyName = child[childPropertyNumber];
                        Assert.AreEqual(row.GetCell(columnNumber).StringCellValue, childPropertyName + index);
                        columnNumber++;
                    }
                }
            }

            for (rowNumber = 1; rowNumber <= sheet.LastRowNum; rowNumber++) {
                row = sheet.GetRow(rowNumber);
                Assert.NotNull(row);
                TestDataEntities.ClientExampleModel currentTestDataRow = simpleTestData.Models[rowNumber - 1];
                Assert.AreEqual(row.GetCell(0).StringCellValue, currentTestDataRow.Title);
                Assert.AreEqual(row.GetCell(1).StringCellValue, currentTestDataRow.RegistrationDate.ToRussianFullString());
                Assert.AreEqual(row.GetCell(2).StringCellValue, currentTestDataRow.Phone);
                Assert.AreEqual(row.GetCell(3).StringCellValue, currentTestDataRow.Inn);
                Assert.AreEqual(row.GetCell(4).StringCellValue, currentTestDataRow.Okato);

                for (int contactNumber = 0; contactNumber < currentTestDataRow.Contacts.Count; contactNumber++) {
                    TestDataEntities.ClientExampleModel.Contact currentContactRow = currentTestDataRow.Contacts[contactNumber];
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
                        currentTestDataRow.Products[productNumber].Amount.ToString());
                }
            }
        }

        [Test]
        public void ExcelStyleSimpleExport() {
            TestDataEntities.TestData simpleTestData = TestDataEntities.CreateSimpleTestData(true);
            TestDataEntities.ClientExampleModel firstTestDataRow = simpleTestData.Models.FirstOrDefault();
            Assert.NotNull(firstTestDataRow);
            TableWriterStyle style = new TableWriterStyle();
            MemoryStream memoryStream = TableWriterSimple.Write(new XlsTableWriterSimple(style), simpleTestData.Models, simpleTestData.ConfigurationBuilder.Value);
            WriteToFile(memoryStream, "TestStyleSimple.xls");
            HSSFWorkbook xlsFile;
            using (FileStream file = new FileStream("TestStyleSimple.xls", FileMode.Open, FileAccess.Read)) {
                xlsFile = new HSSFWorkbook(file);
            }
            ISheet sheet = xlsFile.GetSheetAt(0);

            short[] red = {255, 0, 0};
            short[] green = {0, 255, 0};
            short[] blue = {0, 0, 255};
            
            int rowNumber = 0;
            IRow row = sheet.GetRow(rowNumber);
            Assert.AreEqual(row.Height, 400);
            for (int cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
                IFont cellFont = row.GetCell(cellNumber).CellStyle.GetFont(xlsFile);
                Assert.AreEqual(cellFont.FontName, "Arial");
                Assert.AreEqual(cellFont.FontHeightInPoints, 10);
                Assert.AreEqual(cellFont.Boldweight, (int) FontBoldWeight.Bold);
            }
            
            for (rowNumber = 1; rowNumber < sheet.LastRowNum; rowNumber++) {
                row = sheet.GetRow(rowNumber);
                for (int cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
                    IFont cellFont = row.GetCell(cellNumber).CellStyle.GetFont(xlsFile);

                    if ((rowNumber == 1) & (cellNumber == 2)) {
                        Assert.AreEqual(cellFont.FontName, "Times New Roman");
                        Assert.AreEqual(cellFont.FontHeightInPoints, 14);
                        Assert.AreEqual(cellFont.Boldweight, (int)FontBoldWeight.Bold);
                    }
                    else {
                        Assert.AreEqual(cellFont.FontName, "Arial");
                        Assert.AreEqual(cellFont.FontHeightInPoints, 10);
                        Assert.AreEqual(cellFont.Boldweight, (int) FontBoldWeight.Normal);
                    }
                }
            }
            Assert.IsTrue(EqualsSimpleColors((HSSFColor)sheet.GetRow(1).GetCell(2).CellStyle.FillForegroundColorColor, green));
            Assert.IsTrue(EqualsSimpleColors((HSSFColor)sheet.GetRow(1).GetCell(5).CellStyle.FillForegroundColorColor, blue));
            Assert.IsTrue(EqualsSimpleColors((HSSFColor)sheet.GetRow(2).GetCell(0).CellStyle.FillForegroundColorColor, red));
            Assert.IsTrue(EqualsSimpleColors((HSSFColor)sheet.GetRow(2).GetCell(5).CellStyle.FillForegroundColorColor, blue));
        }

        [Test]
        public void ExcelComplexExport() {
            TestDataEntities.TestData simpleTestData = TestDataEntities.CreateSimpleTestData();
            TestDataEntities.ClientExampleModel firstTestDataRow = simpleTestData.Models.FirstOrDefault();
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
                TestDataEntities.ClientExampleModel currentTestDataRow = simpleTestData.Models[modelNumber];
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
                rowNumber++;
                row = sheet.GetRow(rowNumber);

                int childsQuantity = simpleTestData.ConfigurationBuilder.Value.ChildrenMap.Count;

                for (int childNumber = 0; childNumber < childsQuantity; childNumber++) {
                    switch (childNumber) {
                        case 0: {
                                List<TestDataEntities.ClientExampleModel.Contact> child = currentTestDataRow.Contacts;
                                TestChildHeader(row, childNumber, simpleTestData);
                                rowNumber++;
                                row = sheet.GetRow(rowNumber);
                                for (int childPropertyNumber = 0; childPropertyNumber < child.Count; childPropertyNumber++) {
                                    Assert.AreEqual(row.GetCell(1).StringCellValue, child[childPropertyNumber].Title);
                                    Assert.AreEqual(row.GetCell(2).StringCellValue, child[childPropertyNumber].Email);
                                    rowNumber++;
                                    row = sheet.GetRow(rowNumber);
                                }
                                break;
                            }
                        case 1: {
                                List<TestDataEntities.ClientExampleModel.Contract> child = currentTestDataRow.Contracts;
                                TestChildHeader(row, childNumber, simpleTestData);
                                rowNumber++;
                                row = sheet.GetRow(rowNumber);
                                for (int childPropertyNumber = 0; childPropertyNumber < child.Count; childPropertyNumber++) {
                                    Assert.AreEqual(row.GetCell(1).StringCellValue, child[childPropertyNumber].BeginDate.ToRussianFullString());
                                    Assert.AreEqual(row.GetCell(2).StringCellValue, child[childPropertyNumber].EndDate.ToRussianFullString());
                                    Assert.AreEqual(row.GetCell(3).StringCellValue, child[childPropertyNumber].Status.ToRussianString());
                                    rowNumber++;
                                    row = sheet.GetRow(rowNumber);
                                }
                                break;
                            }
                        case 2: {
                                List<TestDataEntities.ClientExampleModel.Product> child = currentTestDataRow.Products;
                                TestChildHeader(row, childNumber, simpleTestData);
                                rowNumber++;
                                row = sheet.GetRow(rowNumber);
                                for (int childPropertyNumber = 0; childPropertyNumber < child.Count; childPropertyNumber++) {
                                    Assert.AreEqual(row.GetCell(1).StringCellValue, child[childPropertyNumber].Title);
                                    Assert.AreEqual(row.GetCell(2).StringCellValue, child[childPropertyNumber].Amount.ToString());
                                    rowNumber++;
                                    row = sheet.GetRow(rowNumber);
                                }
                                break;
                            }
                    }
                }
            }
        }

        private static void TestChildHeader(IRow row, int childNumber, TestDataEntities.TestData simpleTestData) {
            string childName = simpleTestData.ConfigurationBuilder.Value.ChildrenMap[childNumber].Title;
            List<string> childColumnsNames = simpleTestData.ConfigurationBuilder.Value.ChildrenMap[childNumber].ColumnsMap.Keys.ToList();

            Assert.AreEqual(row.GetCell(0).StringCellValue, childName);
            for (int cellNumber = 1; cellNumber < row.LastCellNum; cellNumber++) {
                int numberProperty = cellNumber - 1;
                Assert.AreEqual(row.GetCell(cellNumber).StringCellValue, childColumnsNames[numberProperty]);
            }
        }

        private bool EqualsSimpleColors(HSSFColor objectColor, short[] color) {
            short[] triplet = objectColor.GetTriplet();
            if ((triplet[0] == color[0]) & (triplet[1] == color[1]) & (triplet[2] == color[2])) {
                return true;
            }
            else return false;
        }

        private static void WriteToFile(MemoryStream ms, string fileName) {
            using (FileStream file = new FileStream(fileName, FileMode.Create, FileAccess.Write)) {
                byte[] bytes = new byte[ms.Length];
                ms.Read(bytes, 0, (int)ms.Length);
                file.Write(bytes, 0, bytes.Length);
                ms.Close();
            }
        }

        public enum FontBoldWeight {
            None = 0,
            Normal = 400,
            Bold = 700,
        }
    }
}