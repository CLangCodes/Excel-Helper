using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;

namespace ExcelService;

/// <summary>
/// Goal: Provide a service to import and export Excel files. It should be able to read and write to the Excel file.
/// It should be able to create a new Excel file if it doesn't exist.
/// It should be able to add a new sheet to the Excel file.
/// It should be able to add a new row to the Excel file.
/// It should be able to add a new column to the Excel file.
/// It should be able to add a new cell to the Excel file.
/// It should be able to add a new formula to the Excel file.
/// It should be able to add a new comment to the Excel file.
/// It should be able to add a new hyperlink to the Excel file.
/// Does a spreadsheetDocument have multiple workbookParts? 
/// Is a workbookPart a worksheet? Or vise versa?
/// What is the difference between a workbookPart and a worksheetPart?
/// What is sheetData? Is that what holds the rows and columns?
/// Does every worksheet have a shared string table?
/// </summary>

public class ExcelService
{
    public static SpreadsheetDocument CreateSpreadsheetWorkbook(string filepath)
    {
        // Create a spreadsheet document by supplying the filepath.
        // By default, AutoSave = true, Editable = true, and Type = xlsx.
        using (SpreadsheetDocument document = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook))
        {
            // Add a WorkbookPart to the document.
            WorkbookPart workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "mySheet" };
            sheets.Append(sheet);
            return document;
        }
    }

    public static SpreadsheetDocument ImportExcelFile(string filePath)
    {
        // We want to open the file in read-write mode.
        using (var document = SpreadsheetDocument.Open(filePath, true))
        {
            return document;
        }
    }

    public static void ExportExcelFile(string filePath, SpreadsheetDocument document)
    {
        if (document.FileOpenAccess != FileAccess.ReadWrite) throw new Exception("The document is not open in read-write mode.");
        document.Save();
    }

    public Row GetRow(SpreadsheetDocument document, string sheetName, int rowIndex)
    {
        Console.WriteLine("Getting row...");
        var sheet = document.WorkbookPart.Workbook.GetFirstChild<Sheet>().Elements<SheetData>().First();
        var row = sheet.Elements<Row>().ElementAt(rowIndex);
        return row;
    }

    public Column GetColumn(SpreadsheetDocument document, string sheetName, int columnIndex)
    {
        Console.WriteLine("Getting column...");
        var sheet = document.WorkbookPart.Workbook.GetFirstChild<Sheet>().Elements<SheetData>().First();
        var column = sheet.Elements<Column>().ElementAt(columnIndex);
        return column;
    }

    public Cell GetCell(SpreadsheetDocument document, string sheetName, int rowIndex, int columnIndex)
    {
        Console.WriteLine("Getting cell...");
        var sheet = document.WorkbookPart.Workbook.GetFirstChild<Sheet>().Elements<SheetData>().First();
        var cell = sheet.Elements<Cell>().ElementAt(columnIndex);
        return cell;
    }
}