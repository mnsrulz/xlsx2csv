using xlsx2csv;
using Shouldly;
namespace xlsx2csv.Tests;

public class UtilityTests
{
    [Fact]
    public void ConvertWithAcceptedSeparators()
    {
        XlsxUtility.Convert(new Options
        {
            InputFileName = "TestData/simple-single-sheet.xlsx",
            Delimiter = '\t'
        });
    }
    
    [Fact]
    public void DefaultOutputFileNameTest()
    {
        XlsxUtility.Convert(new Options
        {
            InputFileName = "TestData/simple-single-sheet.xlsx"
        });
        File.Exists("TestData/simple-single-sheet.csv").ShouldBeTrue();
    }
}