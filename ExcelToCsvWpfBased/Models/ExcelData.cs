using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;

namespace ExcelToCsvWpfBased
{
    internal sealed class ExcelData : IDisposable
    {
        private const string myConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\"{0}\";Extended Properties=Excel 12.0;";
        private OleDbConnection myOleDbConnection;
        private bool myIsNotDisposed = true;

        public void Open(string filename)
        {
            try
            {
                string connectionString = string.Format(myConnectionString, filename);
                this.myOleDbConnection = new OleDbConnection(connectionString);
                myOleDbConnection.Open();
            }
            catch (InvalidOperationException)
            {
                throw;
            }
            catch (OleDbException)
            {
                throw new DatabaseConnectionException();
            }
        }

        public IEnumerable<string> GetSheetNames()
        {
            DataTable schemaTable = this.myOleDbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            foreach (DataRow dataRow in schemaTable.AsEnumerable())
            {
                string coloumnName = dataRow.Field<string>("TABLE_NAME");
                yield return this.GetSheetName(coloumnName);
            }
        }

        public void GetSheetData(StreamWriter output, string sheetName)
        {
            string tableName = this.GetTableName(sheetName);
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [" + tableName + "]", this.myOleDbConnection);

            DataSet ds = new DataSet();
            da.FillSchema(ds, SchemaType.Source, tableName);
            var table = ds.Tables[tableName];
            foreach (DataColumn column in table.Columns)
            {
                column.DataType = typeof(string);
            }

            // Output column names
            output.WriteLine(this.ArrayToCsv(table.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));

            // Output row data
            da.Fill(ds, tableName);
            foreach (var row in table.AsEnumerable())
            {
                output.WriteLine(this.ArrayToCsv(row.ItemArray.Select(o => o.ToString())));
            }
        }

        // There is no standard for 'CSV', so we use the basic rules for RFC 4180
        private string ArrayToCsv(IEnumerable<string> arr)
        {
            List<string> resultValues = new List<string>();
            foreach (string input in arr)
            {
                string output = input.Replace("\"", "\"\"");

                if (input.Contains("\r") || input.Contains("\n") || input.Contains("\"") || input.Contains(","))
                {
                    output = "\"" + output + "\"";
                }
                resultValues.Add(output);

            }

            return string.Join(",", resultValues);
        }

        private string GetSheetName(string tableName)
        {
            string sheetName = tableName;

            if (sheetName.Length >= 2 && sheetName.StartsWith("'") && sheetName.EndsWith("'"))
            {
                sheetName = sheetName.Substring(1, sheetName.Length - 2);
            }

            if (sheetName.EndsWith("$"))
            {
                sheetName = sheetName.Substring(0, sheetName.Length - 1);
            }

            return sheetName;
        }

        private string GetTableName(string sheetName)
        {
            if (sheetName.Contains(" "))
            {
                return string.Format("'{0}$'", sheetName);
            }
            else
            {
                return string.Format("{0}$", sheetName);
            }
        }

        public void Dispose()
        {
            if (myIsNotDisposed)
            {
                myOleDbConnection.Close();
                myOleDbConnection.Dispose();

                myOleDbConnection = null;

                myIsNotDisposed = false;
            }
        }
    }

}
