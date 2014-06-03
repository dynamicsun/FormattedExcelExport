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
            CustomAssert.IsEqualExcelColor((HSSFColor)sheet.GetRow(1).GetCell(5).CellStyle.FillForegroundColorColor, blue);
            CustomAssert.IsEqualExcelColor((HSSFColor)sheet.GetRow(2).GetCell(0).CellStyle.FillForegroundColorColor, red);
            CustomAssert.IsEqualExcelColor((HSSFColor)sheet.GetRow(2).GetCell(5).CellStyle.FillForegroundColorColor, blue);
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

        [Test]
        public void ExcelStyleComplexExport() {
            TestDataEntities.TestData simpleTestData = TestDataEntities.CreateSimpleTestData(true);
            TestDataEntities.ClientExampleModel firstTestDataRow = simpleTestData.Models.FirstOrDefault();
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
            CustomAssert.IsEqualExcelColor((HSSFColor)sheet.GetRow(12).GetCell(1).CellStyle.FillForegroundColorColor, red);
            CustomAssert.IsEqualExcelColor((HSSFColor)sheet.GetRow(14).GetCell(1).CellStyle.FillForegroundColorColor, blue);

            int modelsQuantity = simpleTestData.Models.Count;
            int rowNumber = 0;
            int childsQuantity = 0;
            for (int modelNumber = 0; modelNumber < modelsQuantity; modelNumber++) {
                childsQuantity += 5 + simpleTestData.Models[modelNumber].Contacts.Count + simpleTestData.Models[modelNumber].Contracts.Count + simpleTestData.Models[modelNumber].Products.Count;
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
            string fileName = "TestReflectionSimple.xls";
            WriteToFile(memoryStream, fileName);

            ExcelReflectionSimpleExportTest(models, fileName);
        }

        [Test]
        public void ExcelStyleReflectionSimpleExport() {
            List<TestDataEntities.ClientExampleModel> models = TestDataEntities.CreateSimpleTestDataModels();
            TableWriterStyle style = new TableWriterStyle();
            MemoryStream memoryStream = ReflectionWriterSimple.Write(models, new XlsTableWriterSimple(style), new CultureInfo("ru-Ru"));
            string fileName = "TestReflectionStyleSimple.xls";
            WriteToFile(memoryStream, fileName);

            Assert.NotNull(models.FirstOrDefault());

            HSSFWorkbook xlsFile;
            using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read)) {
                xlsFile = new HSSFWorkbook(file);
            }
            ISheet sheet = xlsFile.GetSheetAt(0);

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
        }
        [Test]
        public void ExcelReflectionComplexExport() {
            List<TestDataEntities.ClientExampleModel> models = TestDataEntities.CreateSimpleTestDataModels();
            TableWriterStyle style = new TableWriterStyle();
            MemoryStream memoryStream = ReflectionWriterComplex.Write(models, new XlsTableWriterComplex(style), new CultureInfo("ru-Ru"));
            string fileName = "TestReflectionComplex.xls";
            WriteToFile(memoryStream, fileName);

            ExcelReflectionComplexExportTest(models, fileName);
        }

        [Test]
        public void ExcelStyleReflectionComplexExport() {
            List<TestDataEntities.ClientExampleModel> models = TestDataEntities.CreateSimpleTestDataModels();
            TableWriterStyle style = new TableWriterStyle();
            MemoryStream memoryStream = ReflectionWriterComplex.Write(models, new XlsTableWriterComplex(style), new CultureInfo("ru-Ru"));
            const string fileName = "TestReflectionStyleComplex.xls";
            WriteToFile(memoryStream, fileName);

            ExcelStyleReflectionComplexExportTest(models, fileName);
        }

        public void ExcelReflectionSimpleExportTest<T>(List<T> models, string fileName) {
            T firstModel = models.FirstOrDefault();
            Assert.NotNull(firstModel);
            List<Type> generalTypes = GeneralTypes;
            List<PropertyInfo> nonEnumerableProperties = firstModel.GetType().GetProperties().Where(x => generalTypes.Contains(x.PropertyType)).ToList();
            List<PropertyInfo> enumerableProperties = firstModel.GetType().GetProperties().Where(x => x.PropertyType.IsGenericType && x.PropertyType.GetGenericTypeDefinition() == typeof(List<>)).ToList();

            HSSFWorkbook xlsFile;
            using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read)) {
                xlsFile = new HSSFWorkbook(file);
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
        public void ExcelReflectionComplexExportTest<T>(List<T> models, string fileName) {
            T firstModel = models.FirstOrDefault();
            Assert.NotNull(firstModel);
            List<Type> generalTypes = GeneralTypes;
            List<PropertyInfo> nonEnumerableProperties = firstModel.GetType().GetProperties().Where(x => generalTypes.Contains(x.PropertyType)).ToList();
            List<PropertyInfo> enumerableProperties = firstModel.GetType().GetProperties().Where(x => x.PropertyType.IsGenericType && x.PropertyType.GetGenericTypeDefinition() == typeof(List<>)).ToList();

            HSSFWorkbook xlsFile;
            using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read)) {
                xlsFile = new HSSFWorkbook(file);
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

        public void ExcelStyleReflectionComplexExportTest<T>(List<T> models, string fileName) {
            Assert.NotNull(models.FirstOrDefault());
            HSSFWorkbook xlsFile;
            using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read)) {
                xlsFile = new HSSFWorkbook(file);
            }
            ISheet sheet = xlsFile.GetSheetAt(0);

            int rowNumber = 0;
            IRow row = sheet.GetRow(rowNumber);
            foreach (T model in models) {
                int cellNumber = 0;
                Assert.AreEqual(row.Height, 400);
                for (cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
                    CustomAssert.IsEqualFont(xlsFile, sheet, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Bold);
                }
                rowNumber++;
                row = sheet.GetRow(rowNumber);
                List<PropertyInfo> enumerableProperties = model.GetType().GetProperties().Where(x => x.PropertyType.IsGenericType && x.PropertyType.GetGenericTypeDefinition() == typeof(List<>)).ToList();
                for (cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
                    CustomAssert.IsEqualFont(xlsFile, sheet, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Normal);
                }
                rowNumber++;
                row = sheet.GetRow(rowNumber);
                foreach (PropertyInfo enumerableProperty in enumerableProperties) {
                    IList childs = (IList)enumerableProperty.GetValue(model);
                    for (cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
                        CustomAssert.IsEqualFont(xlsFile, sheet, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Normal);
                    }
                    rowNumber++;

                    int rowNumberStart = rowNumber;
                    for (rowNumber = rowNumberStart; rowNumber < rowNumberStart + childs.Count; rowNumber++) {
                        row = sheet.GetRow(rowNumber);
                        for (cellNumber = 0; cellNumber < row.LastCellNum; cellNumber++) {
                            CustomAssert.IsEqualFont(xlsFile, sheet, rowNumber, cellNumber, "Arial", 10, (short)FontBoldWeight.Normal);
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
        private static void TestChildHeader(IRow row, int childNumber, TestDataEntities.TestData simpleTestData) {
            string childName = simpleTestData.ConfigurationBuilder.Value.ChildrenMap[childNumber].Title;
            List<string> childColumnsNames = simpleTestData.ConfigurationBuilder.Value.ChildrenMap[childNumber].ColumnsMap.Keys.ToList();

            Assert.AreEqual(row.GetCell(0).StringCellValue, childName);
            for (int cellNumber = 1; cellNumber < row.LastCellNum; cellNumber++) {
                int numberProperty = cellNumber - 1;
                Assert.AreEqual(row.GetCell(cellNumber).StringCellValue, childColumnsNames[numberProperty]);
            }
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