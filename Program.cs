using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using StatementImporter;
using Microsoft.Office.Interop.Excel;

// Needs to outside the try so we can use it in the finally block
Application excelApp = null;
Workbooks excelWorkbooks = null;
Workbook excelWorkbook = null;
Worksheet transWorksheet = null;
ListObject transTable = null;
ListRows transRows = null;
Excel.Range dataRange = null;
List<ListRow> rowsToDispose = null;

var excelFile = @"C:\Users\mark\Downloads\import-test.xlsx";
var csvFile = @"C:\Users\mark\Downloads\25012025_7645.csv";

var rawStatment = File.ReadAllText(csvFile);

var statementLines = StatementFormater.FromatStatment(rawStatment);

try
{
    excelApp = new Application();
    excelApp.Visible = false;

    // In order for the Marshal.ReleaseComObject to work, we to avoid using
    // double dots in any excel object. This creates another com object as
    // a container and since we don't have a reference to it we can't release
    // it, so we have to assign a variable for every excel object, even if we
    // are just using it to get to another excel object.
    excelWorkbooks = excelApp.Workbooks;
    excelWorkbook = excelWorkbooks.Open(
        excelFile, false, false, Type.Missing, Type.Missing, Type.Missing,
        Type.Missing, Type.Missing, Type.Missing, false, Type.Missing, Type.Missing,
        Type.Missing, Type.Missing, Type.Missing);

    // Find sheet...
    transWorksheet = excelWorkbook.Worksheets["Transactions"];
    if (transWorksheet == null)
    {
        throw new Exception("Transactions sheet not found");
    }
    // ...and table
    transTable = transWorksheet.ListObjects["TransactionsTable"];
    if (transTable == null)
    {
        throw new Exception("TransactionTable not found");
    }

    // Remember, every single excel object needs to get it's own variable
    // EVERY SINGLE ONE!
    transRows = transTable.ListRows;
    rowsToDispose = new List<ListRow>();
    foreach(var line in statementLines)
    {
        // We want to add out lines at the top, not the bottom of the list
        // (table) object
        // Mmmm how do I handle this object, it gets created afresh each
        // iteration, but I need to dispose of it to
        var newRow = transRows.Add(1);

        // Gives us an empty row
        dataRange = newRow.Range;
        dataRange.Cells[1].Value = line.Account;
        dataRange.Cells[2].Value = line.Date.ToOADate();
        dataRange.Cells[3].Value = line.Descritpion;
        if (line.Debit != 0m)
        {
            dataRange.Cells[4].Value = line.Debit;
        }
        else
        {
            dataRange.Cells[5].Value = line.Credit;
        }

        // Mark all our range objects for disposal
        rowsToDispose.Add(newRow);
    }

    excelWorkbook.Save();
}
catch (COMException comEx)
{
    Console.WriteLine("Problem with Excel: " + comEx.Message);
    Console.Write(comEx.StackTrace);
}
catch (Exception ex)
{
    Console.WriteLine(ex.Message);
    Console.Write(ex.StackTrace);
}
finally
{
    // We have to tell the OS to release *every* COM object we use!
    if (rowsToDispose != null)
    {
        foreach(var row in rowsToDispose)
        {
            Marshal.ReleaseComObject(row);
        }

        rowsToDispose.Clear();
    }
    if (dataRange != null)
    {
        Marshal.ReleaseComObject(dataRange);
        dataRange = null;
    }
    if (transTable != null)
    {
        Marshal.ReleaseComObject(transTable);
        transTable = null;
    }
    if (transRows != null)
    {
        Marshal.ReleaseComObject(transRows);
        transRows = null;
    }
    if (excelWorkbooks != null)
    {
        Marshal.ReleaseComObject(excelWorkbooks);
        excelWorkbooks = null;
    }
    if (transWorksheet != null)
    {
        Marshal.ReleaseComObject(transWorksheet);
        transWorksheet = null;
    }
    if (excelWorkbook != null)
    {
        // Wait till now to close the workbook since I've had problems
        // getting the books open or closed status, here we don't
        // care, just close it
        excelWorkbook.Close(false, Type.Missing, Type.Missing);
        Marshal.ReleaseComObject(excelWorkbook);
        excelWorkbook = null;
    }
    if (excelApp != null)
    {
        excelApp.Quit();
        Marshal.FinalReleaseComObject(excelApp);
        excelApp = null;
    }
}