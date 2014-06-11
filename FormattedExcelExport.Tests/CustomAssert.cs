using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NUnit.Framework;

namespace FormattedExcelExport.Tests {
    internal static class CustomAssert {
        internal static void IsEqualExcelColor(HSSFColor objectColor, short[] color) {
            short[] triplet = objectColor.GetTriplet();
            if ((triplet[0] == color[0]) & (triplet[1] == color[1]) & (triplet[2] == color[2])) return;
            Assert.Fail("Expected {0}, but was {1}", string.Format("[{0},{1},{2}]", color[0], color[1], color[2]), string.Format("[{0},{1},{2}]", triplet[0], triplet[1], triplet[2]));
        }

        internal static void IsEqualExcelColor(XSSFColor objectColor, short[] color) {
            byte[] triplet = objectColor.GetRgb();
            if ((triplet[0] == color[0]) & (triplet[1] == color[1]) & (triplet[2] == color[2])) return;
            Assert.Fail("Expected {0}, but was {1}", string.Format("[{0},{1},{2}]", color[0], color[1], color[2]), string.Format("[{0},{1},{2}]", triplet[0], triplet[1], triplet[2]));
        }

        internal static void IsEqualFont(IWorkbook xlsFile, ISheet sheet, int rowNumber, int cellNumber, string fontName, short fontSize, short fontBoldWeight) {
            IFont cellFont = sheet.GetRow(rowNumber).GetCell(cellNumber).CellStyle.GetFont(xlsFile);
            if (cellFont.FontName == fontName) {
                if (cellFont.FontHeightInPoints == fontSize) {
                    if (cellFont.Boldweight == fontBoldWeight) {
                    }
                    else Assert.Fail("Expected {0}, but was {1}", (int) fontBoldWeight, cellFont.Boldweight);
                } else Assert.Fail("Expected {0}, but was {1}", fontSize, cellFont.FontHeightInPoints);
            } else Assert.Fail("Expected {0}, but was {1}", fontName, cellFont.FontName);            
        }

        internal static void IsEqualFont(XSSFWorkbook xlsFile, int rowNumber, int cellNumber, string fontName, short fontSize, short fontBoldWeight) {
            //GetFont(IWorkbook) почему-то не работает, поэтому шрифт пришлось доставать в лоб
            XSSFSheet sheet1 = (XSSFSheet) xlsFile.GetSheetAt(0);
            XSSFRow row = (XSSFRow) sheet1.GetRow(rowNumber);
            XSSFCell cell = (XSSFCell) row.GetCell(cellNumber);
            XSSFCellStyle cellStyle = (XSSFCellStyle) cell.CellStyle;
            XSSFFont cellFont = cellStyle.GetFont();
            if (cellFont.FontName == fontName) {
                if (cellFont.FontHeightInPoints == fontSize) {
                    if (cellFont.Boldweight == fontBoldWeight) {
                    } else Assert.Fail("Expected {0}, but was {1}", (int)fontBoldWeight, cellFont.Boldweight);
                } else Assert.Fail("Expected {0}, but was {1}", fontSize, cellFont.FontHeightInPoints);
            } else Assert.Fail("Expected {0}, but was {1}", fontName, cellFont.FontName);
        }
    }
}