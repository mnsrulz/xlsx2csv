using System.Diagnostics;
using CommandLine;
using Sylvan.Data.Csv;
using Sylvan.Data.Excel;

Parser.Default.ParseArguments<Options>(args)
                    .WithParsed<Options>(o =>
                    {
                        long startTime = Stopwatch.GetTimestamp();
                        FileInfo inputFile = new FileInfo(o.InputFileName);
                        if (!inputFile.Exists)
                        {
                            Console.WriteLine("File not exists");
                            return;
                        }

                        if (String.IsNullOrEmpty(o.OutputFileName))
                            o.OutputFileName = Path.ChangeExtension(inputFile.FullName, ".csv");

                        var edr = ExcelDataReader.Create(o.InputFileName, new ExcelDataReaderOptions
                        {
                            GetErrorAsNull = true
                        });
                        do
                        {
                            var sheetName = edr.WorksheetName;
                            using (CsvDataWriter cdw = CsvDataWriter.Create("data-" + sheetName + ".csv"))
                            {
                                cdw.Write(edr);
                            }
                        } while (edr.NextResult());

                        TimeSpan elapsedTime = Stopwatch.GetElapsedTime(startTime);
                        Console.WriteLine($"Converted file in {elapsedTime}");
                    });

class Options
{
    [Value(0, Required = true, MetaName = "InputFile", HelpText = "Input file to be processed.")]
    public string InputFileName { get; set; }

    [Value(2, Required = false, MetaName = "Worksheet", HelpText = "Worksheet name to be processed.")]
    public string WorksheetName { get; set; }

    [Value(1, Required = false, MetaName = "OutputFile", HelpText = "Output file to write data to.")]
    public string OutputFileName { get; set; }


    // [Option("password", Required = false, HelpText = "Password for open xlsx file.")]
    // public string Password { get; set; }

    // [Option("encoding", Required = false, HelpText = "CSV file encoding.", Default = "utf-8")]
    // public string Encoding { get; set; }

    // [Option("separator", Required = false, HelpText = "CSV file separator.")]
    // public string Separator { get; set; }

    // [Option("language", Required = false, HelpText = "CSV file language (culture).")]
    // public string Language { get; set; }

    // [Option("use-quote", Required = false, HelpText = "When use quote (values WhenNeeded, Always, Never).", Default = Quote.WhenNeeded)]
    // public Quote DataQuote { get; set; }

    // [Option("header-quote", Required = false, HelpText = "When use quote in header row (values WhenNeeded, Always, Never).", Default = Quote.WhenNeeded)]
    // public Quote HeaderQuote { get; set; }


    // [Option("silent", Required = false, HelpText = "Set output to silent (no messages).")]
    // public bool Silent { get; set; }

    // [Option("verbose", Required = false, HelpText = "Set output to verbose messages.")]
    // public bool Verbose { get; set; }
}