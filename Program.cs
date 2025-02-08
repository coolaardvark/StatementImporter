using System.Runtime.InteropServices;
using StatementImporter;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;

// Every excel object needs to outside the try so we can use them in
// the finally block
Application excelApp = null;
Workbook excelWorkbook = null;

(string excelFile, string csvFile) = processArgs(args);

var rawStatment = File.ReadAllText(csvFile);
// Convert in to a format we can use in the spreadsheet
var statementLines = StatementFormater.FromatStatment(rawStatment);

try
{
    excelApp = new Application();
    excelApp.Visible = false;

    // I could use the 'recomeneded' way to deal with closing Excel
    // and track every single Excel object I create then dispose and
    // free each one at the end, the trouble is that this *only* works
    // if you are an admin and I want this to work for everyone, so I
    // have to use the brute force approch and force the Excel instance
    // to quit at the end. Feels like a hack, but it is the only way that
    // will work for both admins and normal users.
    excelWorkbook = excelApp.Workbooks.Open(
        excelFile, false, false, Type.Missing, Type.Missing, Type.Missing,
        Type.Missing, Type.Missing, Type.Missing, false, Type.Missing, Type.Missing,
        Type.Missing, Type.Missing, Type.Missing);

    // Find sheet...
    var transWorksheet = excelWorkbook.Worksheets["Transactions"];
    if (transWorksheet == null)
    {
        throw new Exception("Transactions sheet not found");
    }
    // ...and table
    var transTable = transWorksheet.ListObjects["TransactionsTable"];
    if (transTable == null)
    {
        throw new Exception("TransactionTable not found");
    }

    foreach(var line in statementLines)
    {
        // We want to add out lines at the top, not the bottom of the list
        // (table) object
        // Mmmm how do I handle this object, it gets created afresh each
        // iteration, but I need to dispose of it to
        var newRow = transTable.ListRows.Add(1);

        // Gives us an empty row
        var dataRange = newRow.Range;
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
    if (excelWorkbook != null)
    {
        // Wait till now to close the workbook since I've had problems
        // getting the books open or closed status, here we don't
        // care, just close it
        excelWorkbook.Close(false, Type.Missing, Type.Missing);
    }
    if (excelApp != null)
    {
        // The only way for a non-admin user to close the excel
        // app opened via interop is to do a forced kill!
        // Feels very much like a hack, but there you go!
        var excelProcessID = 0;

        GetWindowThreadProcessId(
            new IntPtr(excelApp.Hwnd), ref excelProcessID);
        var excelProcess = Process.GetProcessById(excelProcessID);

        excelProcess.Kill();
    }
}

static (string, string) processArgs(string[] args)
{
    if (args.Length != 2)
    {
        Console.WriteLine("Please pass the paths to the files you want me " +
            "to work on");
        Console.WriteLine("One needs to an excel file, the other a csv file, " +
            "the order isn't important");

        Environment.Exit(1);
    }

    var excelFile = string.Empty;
    var csvFile = string.Empty;

    foreach (var arg in args)
    {
        if (File.Exists(arg))
        {
            if (arg.EndsWith(".xlsx"))
            {
                excelFile = arg;
            }
            else if (arg.EndsWith(".csv"))
            {
                csvFile = arg;
            }
            else
            {
                Console.WriteLine("I need an excel file (.xlsx) and a csv file");
                Environment.Exit(1);
            }
        }
        else
        {
            Console.WriteLine($"File {arg} not found");

            Environment.Exit(1);
        }
    }

    return (excelFile, csvFile);
}

[DllImport("user32.dll", SetLastError = true)]
static extern int GetWindowThreadProcessId(IntPtr hwnd, ref int lpdwProcessId);