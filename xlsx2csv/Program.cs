using System.Diagnostics;
using CommandLine;
using Sylvan.Data.Csv;
using Sylvan.Data.Excel;

Parser.Default.ParseArguments<Options>(args)
                    .WithParsed<Options>(o =>
                    {
                        long startTime = Stopwatch.GetTimestamp();
                        var inputFile = new FileInfo(o.InputFileName);
                        if (!inputFile.Exists) throw new FileNotFoundException("File not exists");

                        if (String.IsNullOrEmpty(o.OutputFileName))
                            o.OutputFileName = Path.ChangeExtension(inputFile.FullName, ".csv");

                        if (char.IsWhiteSpace(o.Delimiter))
                            o.Delimiter = ',';

                        var edr = ExcelDataReader.Create(o.InputFileName, new ExcelDataReaderOptions
                        {
                            GetErrorAsNull = true
                        });
                        if (!string.IsNullOrWhiteSpace(o.SheetName))
                        {
                            if (edr.WorksheetNames.Contains(o.SheetName)) edr.TryOpenWorksheet(o.SheetName);
                            else throw new KeyNotFoundException($"Sheet {o.SheetName} not found in the excel file");
                        }

                        do
                        {
                            // var sheetName = edr.WorksheetName;   //for future implementation
                            using (CsvDataWriter cdw = CsvDataWriter.Create(o.OutputFileName, new CsvDataWriterOptions
                            {
                                Delimiter = o.Delimiter
                            }))
                            {
                                cdw.Write(edr);
                            }
                        } while (edr.NextResult());

                        TimeSpan elapsedTime = Stopwatch.GetElapsedTime(startTime);
                        Console.WriteLine($"Converted file in {elapsedTime}");
                    });

class Options
{
    [Value(0, Required = true, MetaName = "xlsxfile", HelpText = "xlsx file path")]
    public string InputFileName { get; set; }

    [Value(1, Required = false, MetaName = "outfile", HelpText = "output csv file path")]
    public string OutputFileName { get; set; }

    [Option('n', "sheetname", Required = false, HelpText = "Worksheet name to be processed.")]
    public string SheetName { get; set; }


    // [Option("password", Required = false, HelpText = "Password for open xlsx file.")]
    // public string Password { get; set; }

    // [Option("encoding", Required = false, HelpText = "CSV file encoding.", Default = "utf-8")]
    // public string Encoding { get; set; }

    [Option('d', "delimiter", Required = false, HelpText = "CSV file separator.")]
    public char Delimiter { get; set; }

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