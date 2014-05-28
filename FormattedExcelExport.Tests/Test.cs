using System.Linq;
using Extensions;
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
            var confBuilder = new TableConfigurationBuilder<TestDataEntities.ClientExampleModel>("Клиент", new CultureInfo("ru-RU"));
            TableWriterStyle condStyle = new TableWriterStyle();
            condStyle.RegularCell.BackgroundColor = new AdHocCellStyle.Color(255, 0, 0);
            TableWriterStyle condStyle2 = new TableWriterStyle();
            condStyle2.RegularCell.BackgroundColor = new AdHocCellStyle.Color(0, 255, 0);
            TableWriterStyle condStyle3 = new TableWriterStyle();
            condStyle3.RegularChildCell.BackgroundColor = new AdHocCellStyle.Color(0, 0, 255);

            confBuilder.RegisterColumn("Название", x => x.Title, new TableConfigurationBuilder<TestDataEntities.ClientExampleModel>.ConditionTheme(condStyle, x => x.Title == "Вторая компания"));
            confBuilder.RegisterColumn("Дата регистрации", x => x.RegistrationDate);
            confBuilder.RegisterColumn("Телефон", x => x.Phone, new TableConfigurationBuilder<TestDataEntities.ClientExampleModel>.ConditionTheme(condStyle2, x => x.Okato == "OPEEHBSSDD"));
            confBuilder.RegisterColumn("ИНН", x => x.Inn);
            confBuilder.RegisterColumn("Окато", x => x.Okato);

            var contact = confBuilder.RegisterChild("Контакт", x => x.Contacts);
            contact.RegisterColumn("Название", x => x.Title, new TableConfigurationBuilder<TestDataEntities.ClientExampleModel.Contact>.ConditionTheme(condStyle3, x => x.Title.StartsWith("О")));
            contact.RegisterColumn("Email", x => x.Email);

            var contract = confBuilder.RegisterChild("Контракт", x => x.Contracts);
            contract.RegisterColumn("Дата начала", x => x.BeginDate);
            contract.RegisterColumn("Дата окончания", x => x.EndDate);
            contract.RegisterColumn("Статус", x => x.Status, new TableConfigurationBuilder<TestDataEntities.ClientExampleModel.Contract>.ConditionTheme(new TableWriterStyle(), x => true));

            var product = confBuilder.RegisterChild("Продукт", x => x.Products);
            product.RegisterColumn("Наименование", x => x.Title);
            product.RegisterColumn("Количество", x => x.Amount);

            List<TestDataEntities.ClientExampleModel> models = InitializeModels();

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

            HSSFWorkbook xlsFile;
            using (FileStream file = new FileStream("TestSimple.xls", FileMode.Open, FileAccess.Read)) {
                xlsFile = new HSSFWorkbook(file);
            }
            ISheet sheet = xlsFile.GetSheetAt(0);
            for (int rowNumber = 1; rowNumber <= sheet.LastRowNum; rowNumber++) {
                IRow row = sheet.GetRow(rowNumber);
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