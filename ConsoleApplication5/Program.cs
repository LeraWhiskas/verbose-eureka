using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using ExcelLibrary.SpreadSheet;

namespace ConsoleApplication5
{
    class Program
    {
        static void Main(string[] args)

        {
            Workbook wb;
            Worksheet ws;
            string path = @"C: \Users\Lera\Documents\Visual Studio 2015\Projects\ConsoleApplication5\workbook.xls";
            try
            {
                wb = Workbook.Load(path);
                ws = wb.Worksheets.First();
            }
            catch
            {
                Console.WriteLine("File does not exist");
                wb = new Workbook();
                ws = new Worksheet("table");
                ws = FillSheet(ws);
                wb.Worksheets.Add(ws);
                wb.Save(path);
            }

            List<obj> data = new List<obj>();
            for (byte i = 0; i < ws.Cells.Rows.Count; i++)
                data.Add(new obj
                {
                    name = ws.Cells[i, 1].ToString(),
                    age = byte.Parse(ws.Cells[i, 2].ToString())
                });
            data = data.Distinct().ToList();

            string comstring = String.Empty;
            string constring = ConsoleApplication5.Properties.Settings.Default.DBConnectionString;
            SqlConnection connection = new SqlConnection(constring);
            SqlCommand command;

            connection.Open();

            DataTable table = new DataTable();
            try
            {
                command = new SqlCommand("select * from tt", connection);
                command.ExecuteNonQuery();
            }
            catch
            {
                Console.WriteLine("table was not found");
                comstring = "create table tt (" +
                    "id int not null identity primary key," +
                    "name varchar (10)," +
                    "age int);";
                command = new SqlCommand(comstring, connection);
                command.ExecuteNonQuery();
            }

            comstring = BuildString(comstring, data);
            command = new SqlCommand(comstring, connection);
            command.ExecuteNonQuery();
            connection.Close();

            Console.ReadLine();
        }

        static string BuildString(string str, List<obj> list)
        {
            str = "insert into tt (name, age) values";
            foreach (var row in list)
                str += " ('" + row.name + "', " + row.age + "),";
            str = str.Remove(str.Length - 1, 1) + ";";
            return str;
        }

        static Worksheet FillSheet(Worksheet sheet)
        {

            string constring = ConsoleApplication5.Properties.Settings.Default.DBConnectionString;
            SqlConnection connection = new SqlConnection(constring);

            DataTable table = new DataTable();

            connection.Open();
            SqlCommand command = new SqlCommand("select * from tt", connection);
            SqlDataAdapter adapter = new SqlDataAdapter(command);
            adapter.Fill(table);
            connection.Close();

            for (byte i = 0; i < 3; i++)
                for (byte j = 0; j < 3; j++)
                    sheet.Cells[i, j] = new Cell(table.Rows[i][j]);

            return sheet;
        }

        class obj
        {
            public byte age;
            public string name;
        }
    }
}
