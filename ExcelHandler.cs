using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using System.Text.RegularExpressions;

namespace ACRConsole.Services;

/// <summary>
/// A static method to help process Excel files.
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

public class ExcelHandler
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
        if (!File.Exists(filePath))
        {
            throw new ArgumentException($"File not found at {filePath}.");
        }

        return SpreadsheetDocument.Open(filePath, true);
    }

    public static void ExportExcelFile(string filePath, SpreadsheetDocument document)
    {
        if (document.FileOpenAccess != FileAccess.ReadWrite) throw new Exception("The document is not open in read-write mode.");
        document.Save();
    }

    public static Row GetRow(SpreadsheetDocument document, string sheetName, int rowIndex)
    {
        if (document == null) throw new ArgumentException("The document is null");
        var workbookPart = document.WorkbookPart;
        if (workbookPart == null) throw new ArgumentException("The workbookPart is null");

        var sheet = workbookPart!.Workbook!.GetFirstChild<Sheet>()!.Elements<SheetData>().First();
        
        Console.WriteLine("Getting row...");
        var row = sheet.Elements<Row>().ElementAt(rowIndex);
        return row;
    }

    public static Column GetColumn(SpreadsheetDocument document, string sheetName, int columnIndex)
    {
        if (document == null) throw new ArgumentException("The document is null");
        var workbookPart = document.WorkbookPart;
        if (workbookPart == null) throw new ArgumentException("The workbookPart is null");

        var sheet = workbookPart!.Workbook!.GetFirstChild<Sheet>()!.Elements<SheetData>().First();

        Console.WriteLine("Getting column...");
        var column = sheet.Elements<Column>().ElementAt(columnIndex);
        return column;
    }

    public static Cell GetCell(SpreadsheetDocument document, string sheetName, int rowIndex, int columnIndex)
    {
        if (document == null) throw new ArgumentException("The document is null");
        var workbookPart = document.WorkbookPart;
        if (workbookPart == null) throw new ArgumentException("The workbookPart is null");

        var sheet = workbookPart!.Workbook!.GetFirstChild<Sheet>()!.Elements<SheetData>().First();

        Console.WriteLine("Getting cell...");

        var cell = sheet.Elements<Cell>().ElementAt(columnIndex);
        return cell;
    }

    /// <summary>
    /// Get a string[] of files in a folder.
    /// </summary>
    public static string[] GetFilesInFolder(string folderPath)
    {
        try
        {
            // Retrieve all files in the directory as a string array of full paths
            string[] filePaths = Directory.GetFiles(folderPath);
            List<string> fileNames = new List<string>();

            foreach (string filePath in filePaths)
            {
                // To get just the filename without the path, use Path.GetFileName()
                var fileAttributes = File.GetAttributes(filePath);
                if (fileAttributes.HasFlag(FileAttributes.Hidden)) continue;
                string fileName = Path.GetFileName(filePath);
                
                fileNames.Add(fileName);
                // Console.WriteLine(fileName);
            }
            return fileNames.ToArray();
        }
        catch (DirectoryNotFoundException)
        {
            Console.WriteLine("The specified directory does not exist.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
        return [];
    }
    
    public static string GetCellValue(SpreadsheetDocument document, string sheetName, string addressName)
    {
        string? value = null;
        // Open the spreadsheet document for read-only access.
        // Retrieve a reference to the workbook part.
        if (document.WorkbookPart == null) { throw new ArgumentException("document.WorkbookPart is null"); }
        
        WorkbookPart wbPart = document.WorkbookPart;

        // Find the sheet with the supplied name, and then use that 
        // Sheet object to retrieve a reference to the first worksheet.
        if (wbPart == null) throw new ArgumentException("WorkbookPart is null");
        
        var workbook = wbPart.Workbook;
        if (wbPart.Workbook == null) throw new ArgumentException("Workbook is empty");

        Sheet sheet = workbook!.Descendants<Sheet>()!.FirstOrDefault(s => s.Name == sheetName)!;

        // Throw an exception if there is no sheet.
        if (sheet is null || sheet.Id is null)
        {
            throw new ArgumentException("sheetName");
        }

        // Retrieve a reference to the worksheet part.
        WorksheetPart wsPart = (WorksheetPart)wbPart!.GetPartById(sheet.Id!);
        // Use its Worksheet property to get a reference to the cell 
        // whose address matches the address you supplied.
        Cell? theCell = wsPart.Worksheet?.Descendants<Cell>()?.Where(c => c.CellReference == addressName).FirstOrDefault();
        // If the cell does not exist, return an empty string.

        if (theCell is null || theCell.InnerText.Length < 0)
        {
            return string.Empty;
        }

        value = theCell.InnerText;
        // If the cell represents an integer number, you are done. 
        // For dates, this code returns the serialized value that 
        // represents the date. The code handles strings and 
        // Booleans individually. For shared strings, the code 
        // looks up the corresponding value in the shared string 
        // table. For Booleans, the code converts the value into 
        // the words TRUE or FALSE.

        if (theCell.DataType is not null)
        {
            if (theCell.DataType.Value == CellValues.SharedString)
            {
                // For shared strings, look up the value in the
                // shared strings table.
                var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                // If the shared string table is missing, something 
                // is wrong. Return the index that is in
                // the cell. Otherwise, look up the correct text in 
                // the table.
                if (stringTable is not null)
                {
                    value = stringTable!.SharedStringTable!.ElementAt(int.Parse(value)).InnerText;
                }
            }
            else if (theCell.DataType.Value == CellValues.Boolean)
            {
                switch (value)
                {
                    case "0":
                        value = "FALSE";
                        break;
                    default:
                        value = "TRUE";
                        break;
                }
            }
        }
        return value;
    }

    // Given a document name and text, 
    // inserts a new work sheet and writes the text to cell "A1" of the new worksheet.
    public static void InsertText(SpreadsheetDocument document, string sheetName, string text, string cellReference)
    {
        WorkbookPart workbookPart = document.WorkbookPart ?? document.AddWorkbookPart();

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

        // Get or Insert a worksheet.

        WorksheetPart worksheetPart = GetWorksheetPartByName(workbookPart, sheetName) ?? InsertWorksheet(workbookPart);
        
        var col = GetColName(cellReference) ?? string.Empty;
        var row = GetRowIndex(cellReference);

        Cell cell = InsertCellInWorksheet(col, (uint)row!, worksheetPart);

        // Set the value of cell A1.
        cell.CellValue = new CellValue(index.ToString());
        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
        document.Save();
    }

    // Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
    // and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
    public static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
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
    public static WorksheetPart InsertWorksheet(WorkbookPart workbookPart)
    {
        // Add a new worksheet part to the workbook.
        WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        newWorksheetPart.Worksheet = new Worksheet(new SheetData());

        Sheets sheets = workbookPart!.Workbook!.GetFirstChild<Sheets>() ?? workbookPart.Workbook.AppendChild(new Sheets());
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
    public static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
    {
        Worksheet worksheet = worksheetPart!.Worksheet!;
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

    public static int? GetColIndex(string cellReference)
    {
        if (string.IsNullOrEmpty(cellReference))
        {
            return null;
        }

        //remove digits
        string columnReference = Regex.Replace(cellReference.ToUpper(), @"[\d]", string.Empty);

        int columnNumber = -1;
        int mulitplier = 1;

        //working from the end of the letters take the ASCII code less 64 (so A = 1, B =2...etc)
        //then multiply that number by our multiplier (which starts at 1)
        //multiply our multiplier by 26 as there are 26 letters
        foreach (char c in columnReference.ToCharArray().Reverse())
        {
            columnNumber += mulitplier * ((int)c - 64);

            mulitplier = mulitplier * 26;
        }

        //the result is zero based so return columnnumber + 1 for a 1 based answer
        //this will match Excel's COLUMN function
        return columnNumber + 1;
    }

    public static string? GetColName(string cellReference)
    {
        if (string.IsNullOrEmpty(cellReference))
        {
            return null;
        }
        return Regex.Replace(cellReference.ToUpper(), @"[\d]", string.Empty);
    }

    public static int? GetRowIndex(string cellReference)
    {
        if (string.IsNullOrEmpty(cellReference))
        {
            return null;
        }
        try
        {
            string colValue = Regex.Replace(cellReference, "[A-Za-z]", string.Empty);
            int.TryParse(colValue, out int value);
            return value;
        } catch
        {
            Console.WriteLine($"Unable to parse the column int from {cellReference}");
            return null;
        } 
    }

    // Get WorksheetPart by sheet name
    public static WorksheetPart? GetWorksheetPartByName(WorkbookPart workbookPart, string sheetName)
    {
        Sheet sheet = workbookPart!.Workbook!.Descendants<Sheet>()
                                           .FirstOrDefault(s => s.Name == sheetName)!;

        if (sheet == null) return null;

        return (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
    }

    // Insert text into a specific cell
    public static void InsertText(WorksheetPart worksheetPart, string cellReference, string text)
    {
        var worksheet = worksheetPart.Worksheet;
        if (worksheet == null) throw new ArgumentException("Worksheet is null");
        SheetData sheetData = worksheet.GetFirstChild<SheetData>()!;

        // Get row index from cell reference (e.g., "B2" -> 2)
        uint rowIndex = uint.Parse(new string(cellReference.Where(char.IsDigit).ToArray()));

        Row row = sheetData!.Elements<Row>()!.FirstOrDefault(r => r!.RowIndex! == rowIndex)!;
        if (row == null)
        {
            row = new Row() { RowIndex = rowIndex };
            sheetData.Append(row);
        }

        Cell cell = row!.Elements<Cell>()!.FirstOrDefault(c => c!.CellReference!.Value == cellReference)!;
        if (cell == null)
        {
            cell = new Cell() { CellReference = cellReference };
            row.Append(cell);
        }

        // Set cell value as inline string
        cell.DataType = CellValues.InlineString;
        cell.InlineString = new InlineString(new Text(text));
    }
}
