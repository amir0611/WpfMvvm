using System;
using System.ComponentModel;
using System.IO;
using Ookii.CommandLine;

namespace ExcelToCsvConsoleBased
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var parser = new CommandLineParser(typeof(ExtractExcelDataParameters));
            ExtractExcelDataParameters arguments;
            try
            {
                arguments = (ExtractExcelDataParameters)parser.Parse(args);
            }
            catch (CommandLineArgumentException ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine();
                parser.WriteUsageToConsole();
                return;
            }

            if (arguments.File == null)
            {
                Console.WriteLine("Required parameter missing - File");
                return;
            }

            var excelData = new ExcelData(arguments.File);
            excelData.Open();

            if (arguments.List)
            {
                foreach (var sheetName in excelData.GetSheetNames())
                {
                    Console.WriteLine(sheetName);
                }
            }

            if (arguments.All)
            {
                foreach (var sheetName in excelData.GetSheetNames())
                {
                    WriteWorksheetToFile(excelData, sheetName, arguments.OutputPath);
                }
            }
            else if (arguments.Worksheet != null)
            {
                WriteWorksheetToFile(excelData, arguments.Worksheet, arguments.OutputPath);
            }

            excelData.Close();
        }
        private static void WriteWorksheetToFile(ExcelData excelData, string sheetName, string outputPath)
        {
            var fileName = sheetName + ".csv";
            if (!string.IsNullOrEmpty(outputPath))
            {
                fileName = Path.Combine(outputPath, fileName);
            }

            var sw = new StreamWriter(fileName);
            excelData.GetSheetData(sw, sheetName);
            sw.Close();
        }
        private class ExtractExcelDataParameters
        {
            /// <summary>
            /// Path of the source Excel file.
            /// It can be set by:-
            /// a) Name, e.g. "-File value".
            /// b) By position,by specifying "value" as the first positional argument
            /// </summary>
            [Description("Path of the source Excel file")]
            [CommandLineArgument(Position = 0, IsRequired = true)]
            public string File { get; set; }

            [Description("Extract data from a specific worksheet")]
            [CommandLineArgument]
            public string Worksheet { get; set; }

            /// <summary>
            /// Lists all sheets.
            /// It can be set by:-
            /// i)   If supplied on the command line, its value will be true, 
            ///      otherwise it will be false. Eg: "-List" to set it to true. 
            /// ii)  It can be explicitly set the value by using "-List:true" or "-List:false" 
            /// iii) This argument has an alias, so it can also be specified using "-L",
            ///      instead of its regular name.
            /// </summary>
            [Description("List all sheets"), Alias("L")]
            [CommandLineArgument]
            public bool List { get; set; }

            [Description("Read data from all sheets")]
            [CommandLineArgument]
            public bool All { get; set; }

            [CommandLineArgument]
            public string OutputPath { get; set; }
        }
    }
}
