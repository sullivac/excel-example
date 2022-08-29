using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelExample
{
    public class Program
    {
        public static void Main(string[] args)
        {
            using (var spreadsheetDocument = SpreadsheetDocument.Create("example.xlsx", SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = spreadsheetDocument.AddWorkbookPart();

                workbookPart.Workbook = new Workbook();

                var sharedStringTablePart = workbookPart.AddNewPart<SharedStringTablePart>();
                sharedStringTablePart.SharedStringTable = new SharedStringTable();
                sharedStringTablePart.SharedStringTable.AppendChild(new SharedStringItem(new Text("Header 1")));
                sharedStringTablePart.SharedStringTable.AppendChild(new SharedStringItem(new Text("Header 2")));
                sharedStringTablePart.SharedStringTable.AppendChild(new SharedStringItem(new Text("Header 3")));
                sharedStringTablePart.SharedStringTable.AppendChild(new SharedStringItem(new Text("Header 4")));
                sharedStringTablePart.SharedStringTable.AppendChild(new SharedStringItem(new Text("Text Value")));

                workbookPart.Workbook.Sheets = new Sheets();
                // Add sheet definition
                workbookPart.Workbook.Sheets.Append(new Sheet { Id = "rId1", Name = "Example", SheetId = 1 });

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId1");

                var sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                var headerRow = sheetData.AppendChild(new Row());
                headerRow.AppendChild(new Cell { CellValue = new CellValue("0"), DataType = CellValues.SharedString });
                headerRow.AppendChild(new Cell { CellValue = new CellValue("1"), DataType = CellValues.SharedString });
                headerRow.AppendChild(new Cell { CellValue = new CellValue("2"), DataType = CellValues.SharedString });
                headerRow.AppendChild(new Cell { CellValue = new CellValue("3"), DataType = CellValues.SharedString });

                var dataRow = sheetData.AppendChild(new Row());
                dataRow.AppendChild(new Cell { CellValue = new CellValue(10.56m) });
                dataRow.AppendChild(new Cell { CellValue = new CellValue(10) });
                dataRow.AppendChild(new Cell { CellValue = new CellValue(15m) });
                dataRow.AppendChild(new Cell { CellValue = new CellValue("4"), DataType = CellValues.SharedString });

                spreadsheetDocument.Save();
                spreadsheetDocument.Close();
            }
        }
    }
}