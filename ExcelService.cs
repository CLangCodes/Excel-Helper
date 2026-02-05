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

    // Given a document name and text, 
    // inserts a new work sheet and writes the text to cell "A1" of the new worksheet.
    static void InsertText(string docName, string text)
    {
        // Open the document for editing.
        using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
        {
            WorkbookPart workbookPart = spreadSheet.WorkbookPart ?? spreadSheet.AddWorkbookPart();

            // Get the SharedStringTablePart. If it does not exist, create a new one.
            SharedStringTablePart shareStringPart;
            if (workbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                shareStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                shareStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
            }

            // Insert the text into the SharedStringTablePart.
            int index = InsertSharedStringItem(text, shareStringPart);

            // Insert a new worksheet.
            WorksheetPart worksheetPart = InsertWorksheet(workbookPart);

            // Insert cell A1 into the new worksheet.
            Cell cell = InsertCellInWorksheet("A", 1, worksheetPart);

            // Set the value of cell A1.
            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
        }
    }

    // Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
    // and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
    static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
    {
        // If the part does not contain a SharedStringTable, create one.
        shareStringPart.SharedStringTable ??= new SharedStringTable();

        int i = 0;

        // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
        foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
        {
            if (item.InnerText == text)
            {
                return i;
            }

            i++;
        }

        // The text does not exist in the part. Create the SharedStringItem and return its index.
        shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));

        return i;
    }

    // Given a WorkbookPart, inserts a new worksheet.
    static WorksheetPart InsertWorksheet(WorkbookPart workbookPart)
    {
        // Add a new worksheet part to the workbook.
        WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        newWorksheetPart.Worksheet = new Worksheet(new SheetData());

        Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>() ?? workbookPart.Workbook.AppendChild(new Sheets());
        string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

        // Get a unique ID for the new sheet.
        uint sheetId = 1;
        if (sheets.Elements<Sheet>().Count() > 0)
        {
            sheetId = sheets.Elements<Sheet>().Select<Sheet, uint>(s =>
            {
                if (s.SheetId is not null && s.SheetId.HasValue)
                {
                    return s.SheetId.Value;
                }

                return 0;
            }).Max() + 1;
        }

        string sheetName = "Sheet" + sheetId;

        // Append the new worksheet and associate it with the workbook.
        Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
        sheets.Append(sheet);

        return newWorksheetPart;
    }

    // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
    // If the cell already exists, returns it. 
    static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
    {
        Worksheet worksheet = worksheetPart.Worksheet;
        SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
        string cellReference = columnName + rowIndex;

        // If the worksheet does not contain a row with the specified row index, insert one.
        Row row;

        if (sheetData?.Elements<Row>().Where(r => r.RowIndex is not null && r.RowIndex == rowIndex).Count() != 0)
        {
            row = sheetData!.Elements<Row>().Where(r => r.RowIndex is not null && r.RowIndex == rowIndex).First();
        }
        else
        {
            row = new Row() { RowIndex = rowIndex };
            sheetData.Append(row);
        }

        // If there is not a cell with the specified column name, insert one.  
        if (row.Elements<Cell>().Where(c => c.CellReference is not null && c.CellReference.Value == columnName + rowIndex).Count() > 0)
        {
            return row.Elements<Cell>().Where(c => c.CellReference is not null && c.CellReference.Value == cellReference).First();
        }
        else
        {
            // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Cell? refCell = null;

            foreach (Cell cell in row.Elements<Cell>())
            {
                if (string.Compare(cell.CellReference?.Value, cellReference, true) > 0)
                {
                    refCell = cell;
                    break;
                }
            }

            Cell newCell = new Cell() { CellReference = cellReference };
            row.InsertBefore(newCell, refCell);

            return newCell;
        }
    }
}