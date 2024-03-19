using ExcelToWordTemplateGen.Generator.Handlers.Word;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;
using System.Data;

namespace ExcelToWordTemplateGen.Tests.Handlers;

public class WordHandlerTests
{
    private readonly WordHandler _wordHandler;
    private readonly DataTable _dtTest;

    public WordHandlerTests()
    {
        var options = Options.Create(new OutputOptions
        {

            DynamicFieldNames = "Firstname;Lastname",
            DynamicFileNameDelimiter = "_",
            StaticFileNameStart = ""
        });

        ILogger<WordHandler> nullLogger = new NullLogger<WordHandler>();
        _wordHandler = new WordHandler(nullLogger, options);
        _dtTest = new DataTable();
        _dtTest.Columns.Add("FIRSTNAME");
        _dtTest.Columns.Add("LASTNAME");
        _dtTest.Rows.Add("value1", "value2");
    }

    [Fact]
    public void SuffixFileNamePart_Returns_ExactFileNameParts()
    {
        string definition = "firstname;lastname";
        string separator = "_";

        string expectedResult = "value1_value2";
        var result = _wordHandler.GetSuffixFileNamePart(definition, separator, _dtTest, _dtTest.Rows[0]);

        Assert.Equal(expectedResult, result);
    }


    [Fact]
    public void SuffixFileNamePart_Returns_ExactFileNamePartsWithoutSeparator()
    {
        string definition = "firstname;lastname";
        string separator = "";

        string expectedResult = "value1value2";
        var result = _wordHandler.GetSuffixFileNamePart(definition, separator, _dtTest, _dtTest.Rows[0]);

        Assert.Equal(expectedResult, result);
    }

    [Fact]
    public void SuffixFileNamePart_Returns_OnlyFirstValueFromColumns()
    {
        string definition = "firstname;lastname2";
        string separator = "_";

        string expectedResult = "value1";
        var result = _wordHandler.GetSuffixFileNamePart(definition, separator, _dtTest, _dtTest.Rows[0]);

        Assert.Equal(expectedResult, result);
    }

    [Fact]
    public void SuffixFileNamePart_Returns_OnlySecondValueFromColumns()
    {
        string definition = "firstname2;lastname";
        string separator = "_";

        string expectedResult = "value2";
        var result = _wordHandler.GetSuffixFileNamePart(definition, separator, _dtTest, _dtTest.Rows[0]);

        Assert.Equal(expectedResult, result);
    }

    [Fact]
    public void SuffixFileNamePart_Returns_EmptyString()
    {
        string definition = "firstname2;lastname2";
        string separator = "_";

        var result = _wordHandler.GetSuffixFileNamePart(definition, separator, _dtTest, _dtTest.Rows[0]);

        Assert.Empty(result);
    }
}
