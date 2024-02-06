using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;

public partial class Program
{
    //static string OriginalFile = @"C:\Users\assad\vscode\stc2excelapp\YourProjectName\stc2excel.xlsx";
    //static string DidNotScan = @"C:\Users\assad\vscode\stc2excelapp\YourProjectName\DidNotScan.csv";
    //static string CheckedInLate = @"C:\Users\assad\vscode\stc2excelapp\YourProjectName\CheckedInLate.csv";
    //static string OnTime = @"C:\Users\assad\vscode\stc2excelapp\YourProjectName\OnTime.csv";

    static string OriginalFile = @"./excel/stc2excel.xlsx";
    static string DidNotScan = @"./excel/DidNotScan.csv";
    static string CheckedInLate = @"./excel/CheckedInLate.csv";
    static string OnTime = @"./excel/OnTime.csv";

        static void Main()
    {
        try
        {
            var connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={OriginalFile};Extended Properties='Excel 12.0 Xml;HDR=YES;'";
            //schema(connectionString);
            GenerateAttendanceTables(connectionString);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
   
    static void GenerateAttendanceTables(string connectionString)
    {
        // Query strings to select data from the sheets
        var queryClassList = "SELECT ID,Name,Group FROM [Sheet1$]";
        var queryAttendanceSheet = "SELECT ID,Date FROM [Sheet2$]";

        var classListTable = new DataTable();
        var attendanceListTable = new DataTable();

        using (var conn = new OleDbConnection(connectionString))
        {
            conn.Open();
            using (var cmd = new OleDbCommand(queryClassList, conn))
            using (var adapter = new OleDbDataAdapter(cmd))
            {
                adapter.Fill(classListTable);
            }

            using (var cmd = new OleDbCommand(queryAttendanceSheet, conn))
            using (var adapter = new OleDbDataAdapter(cmd))
            {
                adapter.Fill(attendanceListTable);
            }
        }

        // Perform left join on classListTable and attendanceListTable
        var query = from t1 in classListTable.AsEnumerable()
                    join t2 in attendanceListTable.AsEnumerable()
                    on t1.Field<double>("ID") equals t2.Field<double>("ID") into gj
                    from subpet in gj.DefaultIfEmpty()
                    select new
                    {
                        ID = t1.Field<double>("ID"),
                        Name = t1.Field<string>("Name"),
                        Group = t1.Field<string>("Group"),
                        Date = subpet == null ? (DateTime?)null : subpet.Field<DateTime>("Date")
                    };

        // Filter based on conditions
        var didNotScan = query.Where(x => x.Date == null).ToList();
        var checkedInLate = query.Where(x => x.Date != null && IsLate(x.Date.Value)).ToList(); 
        var onTime = query.Where(x => x.Date != null && !IsLate(x.Date.Value)).ToList();

        WriteToCsv(didNotScan, DidNotScan);
        WriteToCsv(checkedInLate, CheckedInLate);
        WriteToCsv(onTime, OnTime);

    }

    static bool IsLate(DateTime checkInTime)
    {
        DateTime datePart = checkInTime.Date; 
        DateTime tenAMDateTime = datePart.AddHours(10).AddMinutes(10);
        DateTime noonDateTime = datePart.AddHours(11).AddMinutes(50);
        
        return !(checkInTime <tenAMDateTime || checkInTime >noonDateTime);
    }

    static void WriteToCsv<T>(IEnumerable<T> items, string path)
    {
        using (var writer = new StreamWriter(path))
        {
            
            if (items.Any())
            {
                // Write the header row
                var header = string.Join(",", typeof(T).GetProperties().Select(p => p.Name));
                writer.WriteLine(header);

                // Write the data rows
                foreach (var item in items)
                {
                    var row = string.Join(",", typeof(T).GetProperties().Select(p => p.GetValue(item, null)));
                    writer.WriteLine(row);
                }
            }
        }
    }

    static void schema(string connectionString)
    {
        using (var conn = new OleDbConnection(connectionString))
        {
            conn.Open();

            // Retrieve the schema information for tables (sheets)
            DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            if (schemaTable != null)
            {
                foreach (DataRow row in schemaTable.Rows)
                {
                    // Each row is a table (sheet) in the workbook
                    string sheetName = row["TABLE_NAME"].ToString();
                    Console.WriteLine(sheetName);
                }
            }

        }
    }
}



