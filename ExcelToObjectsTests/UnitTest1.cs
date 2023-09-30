using ExcelToObjects;
// ReSharper disable ClassNeverInstantiated.Local
// ReSharper disable UnassignedGetOnlyAutoProperty

namespace ExcelToObjectsTests;

public class Tests
{
    private string _testFilename = null!;

    [SetUp]
    public void Setup()
    {
        _testFilename = Path.Combine(Environment.CurrentDirectory, @"TestFiles\FirstTest.xlsx");
    }

    [Worksheet(Name = "Sheet1")]
    private class IdentifyWorkSheetButNoColumns { }
    
    [Test]
    public void Given_AttributeWithWorksheetName_Then_TheWorksheetCanBeIdentified()
    {
        var result = ExcelToObjects.ExcelToObjects.ReadData<IdentifyWorkSheetButNoColumns>(_testFilename);
        Assert.That(result.IsValid, Is.True);
    }

    private class WorksheetByClassName { }
    
    [Test]
    public void Given_AttributeWithNoWorksheetName_Then_TheWorksheetCanBeIdentifiedFromTheClassname()
    {
        var result = ExcelToObjects.ExcelToObjects.ReadData<WorksheetByClassName>(_testFilename);
        Assert.That(result.IsValid, Is.True);
    }

    [Worksheet(Name = "This does not exist")]
    private class NoWorksheetFoundTest { }
    
    [Test]
    public void Given_AttributeWithWorksheetThatDoesNotExist_Then_ReturnNotValid()
    {
        var result = ExcelToObjects.ExcelToObjects.ReadData<NoWorksheetFoundTest>(_testFilename);
        Assert.That(result.IsValid, Is.False);
        Assert.That(result.ValidationProblems.Count, Is.EqualTo(1));
        Assert.That(result.ValidationProblems[0].Message, Is.EqualTo("The worksheet could not be found with the name 'This does not exist'."));
    }

    [Worksheet(Name = "NoHeadings")]
    private class OneColumnWithExistingName
    {
        [Column] public string? FirstColumn { get; init; }
    }
    
    [Test]
    public void Given_WorksheetWithNoHeadingsAndColumnIdentifiedByName_Then_TheDataIsMapped()
    {
        var result = ExcelToObjects.ExcelToObjects.ReadData<OneColumnWithExistingName>(_testFilename);
        Assert.That(result.IsValid, Is.True);
        Assert.That(result.Data.Count, Is.EqualTo(3));
        Assert.That(result.Data[0].FirstColumn, Is.EqualTo("one"));
        Assert.That(result.Data[1].FirstColumn, Is.EqualTo("two"));
        Assert.That(result.Data[2].FirstColumn, Is.EqualTo("three"));
    }
    
    [Worksheet(Name = "Sheet1")]
    public class DoubleColumnWorksheet
    {
        [Column(Name = "BB")]
        public double DoubleColumn { get; init; }
    }
    
    [Test]
    public void Given_ColumnDefinitionOfDouble_Then_TheDataIsMappedCorrectly()
    {
        var result = ExcelToObjects.ExcelToObjects.ReadData<DoubleColumnWorksheet>(_testFilename);
        Assert.That(result.IsValid, Is.True);
        Assert.That(result.Data.Count, Is.EqualTo(3));
        Assert.That(result.Data[0].DoubleColumn, Is.EqualTo(1.24d));
        Assert.That(result.Data[1].DoubleColumn, Is.EqualTo(2.34d));
        Assert.That(result.Data[2].DoubleColumn, Is.EqualTo(3.45d));
    }   
}