using System.Collections.Generic;
using System.Linq;
using Extensions;
using FormattedExcelExport.Configuaration;
using NPOI.HSSF.UserModel;
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
                Assert.AreNotEqual(row, null);
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

        //[Test]
        //public void ExcelStyleExport() {
        //    MemoryStream ms;
        //    TableConfigurationBuilder<TestDataEntities.ClientExampleModel> confBuilder;
        //    List<TestDataEntities.ClientExampleModel> models = CreateTestExample(out ms, out confBuilder);
        //    TableWriterStyle style = new TableWriterStyle();
        //    ms = TableWriterComplex.Write(new XlsTableWriterComplex(style), models, confBuilder.Value);
        //    WriteToFile(ms, "TestComplex.xls");


        //}

        //protected ICellStyle ConvertToNpoiStyle(AdHocCellStyle adHocCellStyle, HSSFWorkbook workbook) {
        //    IFont cellFont = workbook.CreateFont();

        //    cellFont.FontName = adHocCellStyle.FontName;
        //    cellFont.FontHeightInPoints = adHocCellStyle.FontHeightInPoints;
        //    cellFont.IsItalic = adHocCellStyle.Italic;
        //    cellFont.Underline = adHocCellStyle.Underline ? FontUnderlineType.Single : FontUnderlineType.None;
        //    cellFont.Boldweight = (short)adHocCellStyle.BoldWeight;

        //    HSSFPalette palette = workbook.GetCustomPalette();
        //    HSSFColor similarColor = palette.FindSimilarColor(adHocCellStyle.FontColor.Red, adHocCellStyle.FontColor.Green, adHocCellStyle.FontColor.Blue);
        //    cellFont.Color = similarColor.GetIndex();

        //    ICellStyle cellStyle = workbook.CreateCellStyle();
        //    cellStyle.SetFont(cellFont);

        //    if (adHocCellStyle.BackgroundColor != null) {
        //        similarColor = palette.FindSimilarColor(adHocCellStyle.BackgroundColor.Red, adHocCellStyle.BackgroundColor.Green, adHocCellStyle.BackgroundColor.Blue);
        //        cellStyle.FillForegroundColor = similarColor.GetIndex();
        //        cellStyle.FillPattern = FillPattern.SolidForeground;
        //    }
        //    return cellStyle;
        //}

        private static void WriteToFile(MemoryStream ms, string fileName) {
            using (FileStream file = new FileStream(fileName, FileMode.Create, FileAccess.Write)) {
                byte[] bytes = new byte[ms.Length];
                ms.Read(bytes, 0, (int)ms.Length);
                file.Write(bytes, 0, bytes.Length);
                ms.Close();
            }
        }
    }
}