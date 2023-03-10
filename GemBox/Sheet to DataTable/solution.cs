using System;
using System.Data;
using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = ExcelFile.Load("SimpleTemplate.xlsx");

        // Create DataTable with specified columns.
        var dataTable = new DataTable();
        dataTable.Columns.Add("First_Column", typeof(string));
        dataTable.Columns.Add("Second_Column", typeof(string));
        dataTable.Columns.Add("Third_Column", typeof(int));
        dataTable.Columns.Add("Fourth_Column", typeof(double));

        // Select the first worksheet from the file.
        var worksheet = workbook.Worksheets[0];

        // Extract the data from an Excel worksheet to the DataTable.
        var options = new ExtractToDataTableOptions(0, 0, 20);
        options.ExcelCellToDataTableCellConverting += (sender, e) =>
        {
            if (!e.IsDataTableValueValid)
            {
                // Convert ExcelCell value to string.
                if (e.DataTableColumnType == typeof(string))
                    e.DataTableValue = e.ExcelCell.Value?.ToString();
                else
                    e.DataTableValue = DBNull.Value;
            }
        };
        worksheet.ExtractToDataTable(dataTable, options);

        // Write DataTable columns.
        foreach (DataColumn column in dataTable.Columns)
            Console.Write(column.ColumnName.PadRight(20));
        Console.WriteLine();
        foreach (DataColumn column in dataTable.Columns)
            Console.Write($"[{column.DataType}]".PadRight(20));
        Console.WriteLine();
        foreach (DataColumn column in dataTable.Columns)
            Console.Write(new string('-', column.ColumnName.Length).PadRight(20));
        Console.WriteLine();

        // Write DataTable rows.
        foreach (DataRow row in dataTable.Rows)
        {
            foreach (object item in row.ItemArray)
            {
                string value = item.ToString();
                value = value.Length > 20 ? value.Remove(19) + "???" : value;
                Console.Write(value.PadRight(20));
            }
            Console.WriteLine();
        }
    }
}
