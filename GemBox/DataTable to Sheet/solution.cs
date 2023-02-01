using System.Data;
using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("DataTable to Sheet");

        var dataTable = new DataTable();

        dataTable.Columns.Add("ID", typeof(int));
        dataTable.Columns.Add("FirstName", typeof(string));
        dataTable.Columns.Add("LastName", typeof(string));

        dataTable.Rows.Add(new object[] { 001, "Shubham", "Headache" });
        dataTable.Rows.Add(new object[] { 002, "Saloni", "Heater" });
       
        worksheet.Cells[0, 0].Value = "DataTable insert example:";

        // Insert DataTable to an Excel worksheet.
        worksheet.InsertDataTable(dataTable,
            new InsertDataTableOptions()
            {
                ColumnHeaders = true,
                StartRow = 2
            });

        workbook.Save("DataTable to Sheet.xlsx");
    }
}
