using ExcelUtilities;
// ReSharper disable ClassNeverInstantiated.Local
// ReSharper disable UnassignedGetOnlyAutoProperty
// ReSharper disable UnusedAutoPropertyAccessor.Local

namespace ExcelToObjectsTests;

public class ExcelToObjectsTests
{
    private string _testFilename = null!;

    [SetUp]
    public void Setup()
    {
        _testFilename = Path.Combine(Environment.CurrentDirectory, @"TestFiles\FirstTest.xlsx");
    }

    [Worksheet(Name = "EmptyWorksheet")]
    private class IdentifyWorkSheetButNoColumns { }
    
    [Test]
    public void Given_AttributeWithWorksheetName_Then_TheWorksheetCanBeIdentified()
    {
        var result = ExcelToObjects.ReadData<IdentifyWorkSheetButNoColumns>(_testFilename);
        Assert.That(result.IsValid, Is.True, $"First validation problem: {result.ValidationProblems.FirstOrDefault()?.Message}");
    }

    private class WorksheetByClassName { }
    
    [Test]
    public void Given_AttributeWithNoWorksheetName_Then_TheWorksheetCanBeIdentifiedFromTheClassname()
    {
        var result = ExcelToObjects.ReadData<WorksheetByClassName>(_testFilename);
        Assert.That(result.IsValid, Is.True, $"First validation problem: {result.ValidationProblems.FirstOrDefault()?.Message}");
    }

    [Worksheet(Name = "This does not exist")]
    private class NoWorksheetFoundTest { }
    
    [Test]
    public void Given_AttributeWithWorksheetThatDoesNotExist_Then_ReturnNotValid()
    {
        var result = ExcelToObjects.ReadData<NoWorksheetFoundTest>(_testFilename);
        Assert.That(result.IsValid, Is.False);
        Assert.That(result.ValidationProblems.Count, Is.EqualTo(1));
        Assert.That(result.ValidationProblems[0].Message, Is.EqualTo("The worksheet could not be found with the name 'This does not exist'."));
    }

    [Worksheet(Name = "NoHeadings")]
    private class OneColumnWithExistingName
    {
        [Column(Name = "B")] public string? ColumnData { get; init; }
    }
    
    [Test]
    public void Given_WorksheetWithNoHeadingsAndColumnIdentifiedByName_Then_TheDataIsMapped()
    {
        var result = ExcelToObjects.ReadData<OneColumnWithExistingName>(_testFilename);
        Assert.That(result.IsValid, Is.True, $"First validation problem: {result.ValidationProblems.FirstOrDefault()?.Message}");
        Assert.That(result.Data.Count, Is.EqualTo(1));
        Assert.That(result.Data[0].ColumnData, Is.EqualTo("find me"));
    }
    
    [Worksheet(Name = "TypeTests")]
    private class DoubleColumnWorksheet
    {
        [Column(Name = "B")] public double DoubleColumn { get; init; }
        [Column(Name = "B")] public double? NullableDoubleColumn { get; init; }
    }
    
    [Test]
    public void Given_ColumnDefinitionOfDouble_Then_TheDataIsMappedCorrectly()
    {
        var result = ExcelToObjects.ReadData<DoubleColumnWorksheet>(_testFilename);
        Assert.That(result.IsValid, Is.True, $"First validation problem: {result.ValidationProblems.FirstOrDefault()?.Message}");
        Assert.That(result.Data[0].DoubleColumn, Is.EqualTo(1.23d));
        Assert.That(result.Data[0].NullableDoubleColumn, Is.EqualTo(1.23d));
    }
    
    [Worksheet(Name = "TypeTests")]
    private class DateOnlyColumnWorksheet
    {
        [Column(Name = "C")] public DateOnly DateOnlyColumn { get; init; }
        [Column(Name = "C")] public DateOnly? NullableDateOnlyColumn { get; init; }
    }
    
    [Test]
    public void Given_ColumnDefinitionOfDateOnly_Then_TheDataIsMappedCorrectly()
    {
        var result = ExcelToObjects.ReadData<DateOnlyColumnWorksheet>(_testFilename);
        Assert.That(result.IsValid, Is.True, $"First validation problem: {result.ValidationProblems.FirstOrDefault()?.Message}");
        Assert.That(result.Data[0].DateOnlyColumn, Is.EqualTo(new DateOnly(2020, 09, 01)));
        Assert.That(result.Data[0].NullableDateOnlyColumn, Is.EqualTo(new DateOnly(2020, 09, 01)));
    }
    
    [Worksheet(Name = "TypeTests")]
    private class DateTimeColumnWorksheet
    {
        [Column(Name = "E")] public DateTime DateTimeColumn { get; init; }
        [Column(Name = "E")] public DateTime? NullableDateTimeColumn { get; init; }
    }
    
    [Test]
    public void Given_ColumnDefinitionOfDateTime_Then_TheDataIsMappedCorrectly()
    {
        var result = ExcelToObjects.ReadData<DateTimeColumnWorksheet>(_testFilename);
        Assert.That(result.IsValid, Is.True, $"First validation problem: {result.ValidationProblems.FirstOrDefault()?.Message}");
        Assert.That(result.Data[0].DateTimeColumn, Is.EqualTo(new DateTime(2021, 04, 02, 10, 45, 00)));
        Assert.That(result.Data[0].NullableDateTimeColumn, Is.EqualTo(new DateTime(2021, 04, 02, 10, 45, 00)));
    }
    
    [Worksheet(Name = "TypeTests")]
    private class TimeOnlyColumnWorksheet
    {
        [Column(Name = "F")] public TimeOnly TimeOnlyColumn { get; init; }
        [Column(Name = "F")] public TimeOnly? NullableTimeOnlyColumn { get; init; }
    }
    
    [Test]
    public void Given_ColumnDefinitionOfTimeOnly_Then_TheDataIsMappedCorrectly()
    {
        var result = ExcelToObjects.ReadData<TimeOnlyColumnWorksheet>(_testFilename);
        Assert.That(result.IsValid, Is.True, $"First validation problem: {result.ValidationProblems.FirstOrDefault()?.Message}");
        Assert.That(result.Data[0].TimeOnlyColumn, Is.EqualTo(new TimeOnly(23, 14, 00)));
        Assert.That(result.Data[0].NullableTimeOnlyColumn, Is.EqualTo(new TimeOnly(23, 14, 00)));
    }
    
    [Worksheet(Name = "TypeTests")]
    private class AccountingColumnWorksheet
    {
        [Column(Name = "D")] public double AccountingColumn { get; init; }
        [Column(Name = "D")] public double? NullableAccountingColumn { get; init; }
    }
    
    [Test]
    public void Given_ColumnDefinitionOfAccounting_Then_TheDataIsMappedCorrectly()
    {
        var result = ExcelToObjects.ReadData<AccountingColumnWorksheet>(_testFilename);
        Assert.That(result.IsValid, Is.True, $"First validation problem: {result.ValidationProblems.FirstOrDefault()?.Message}");
        Assert.That(result.Data[0].AccountingColumn, Is.EqualTo(100.00d));
        Assert.That(result.Data[0].NullableAccountingColumn, Is.EqualTo(100.00d));
    } 
    
    [Worksheet(Name = "TypeTests")]
    private class CurrencyColumnWorksheet
    {
        [Column(Name = "G")] public double CurrencyColumn { get; init; }
        [Column(Name = "G")] public double? NullableCurrencyColumn { get; init; }
    }
    
    [Test]
    public void Given_ColumnDefinitionOfCurrency_Then_TheDataIsMappedCorrectly()
    {
        var result = ExcelToObjects.ReadData<CurrencyColumnWorksheet>(_testFilename);
        Assert.That(result.IsValid, Is.True, $"First validation problem: {result.ValidationProblems.FirstOrDefault()?.Message}");
        Assert.That(result.Data[0].CurrencyColumn, Is.EqualTo(12.23d));
        Assert.That(result.Data[0].NullableCurrencyColumn, Is.EqualTo(12.23d));
    }

    [Worksheet(Name = "WithHeadings", HasHeadings = true)]
    private class WorksheetWithHeadings
    {
        [Column(Heading = "First Column")] public string? FirstColumn { get; init; }
        [Column(Heading = "Second Column")] public double SecondColumn { get; init; }
        [Column(Heading = "Third Column")] public DateOnly ThirdColumn { get; init; }
        [Column(Heading = "Fourth Column")] public double FourthColumn { get; init; }
    }
    
    [Test]
    public void Given_ColumnsWithHeadings_Then_DataIsMappedCorrectly()
    {
        var result = ExcelToObjects.ReadData<WorksheetWithHeadings>(_testFilename);
        Assert.That(result.IsValid, Is.True, $"First validation problem: {result.ValidationProblems.FirstOrDefault()?.Message}");
        Assert.That(result.Data[0].FirstColumn, Is.EqualTo("one"));
        Assert.That(result.Data[0].SecondColumn, Is.EqualTo(1.23d));
        Assert.That(result.Data[0].ThirdColumn, Is.EqualTo(new DateOnly(2020, 09, 01)));
        Assert.That(result.Data[0].FourthColumn, Is.EqualTo(100.00d));
    }
    
    [Worksheet(Name = "WithHeadings", HasHeadings = true)]
    private class WorksheetWithOptionalColumn
    {
        [Column(Heading = "An optional column", Optional = true)] public string? OptionalColumn { get; init; }
    }
    
    [Test]
    public void Given_AColumnMarkedAsOptionalThatDoesNotExist_Then_ItWillNotBePopulated()
    {
        var result = ExcelToObjects.ReadData<WorksheetWithOptionalColumn>(_testFilename);
        Assert.That(result.IsValid, Is.True, $"First validation problem: {result.ValidationProblems.FirstOrDefault()?.Message}");
        Assert.That(result.Data[0].OptionalColumn, Is.Null);
    }

    [Worksheet(Name = "HeadingsOnRowThree", HasHeadings = true, HeadingsOnRow = 3)]
    private class WorksheetWithHeadingsOnRowThree
    {
        [Column(Heading = "First Column")] public double FirstColumn { get; init; }
    }

    [Test]
    public void Given_HeadingsStartOnRowThree_Then_DateIsReadFromRowFour()
    {
        var result = ExcelToObjects.ReadData<WorksheetWithHeadingsOnRowThree>(_testFilename);
        Assert.That(result.IsValid, Is.True, $"First validation problem: {result.ValidationProblems.FirstOrDefault()?.Message}");
        Assert.That(result.Data[0].FirstColumn, Is.EqualTo(234));
    }

    [Worksheet(Name = "WithBlankRows")]
    private class WorksheetWithBlankRowsAndRequiredProperty
    {
        // ReSharper disable UnusedMember.Local
        [Column(Optional = true)] public double? FirstColumn { get; init; }
        [Column] public string? SecondColumn { get; init; }
        // ReSharper restore UnusedMember.Local
    }

    [Test]
    public void Given_WorksheetWithBlankRowsAndRequiredProperties_Then_AValidationProblemIsReturned()
    {
        var result = ExcelToObjects.ReadData<WorksheetWithBlankRowsAndRequiredProperty>(_testFilename);
        Assert.That(result.IsValid, Is.False);
    }
    
    [Worksheet(Name = "WithBlankRows", SkipBlankRows = true)]
    private class WorksheetWithBlankRowsSkipped
    {
        [Column] public double? FirstColumn { get; init; }
        [Column] public string? SecondColumn { get; init; }
    }
    
    [Test]
    public void Given_WorksheetWithBlankRowsAndSkipEnabled_Then_TheDataIsReturnedWithoutTheBlankRows()
    {
        var result = ExcelToObjects.ReadData<WorksheetWithBlankRowsSkipped>(_testFilename);
        Assert.That(result.IsValid, Is.True, $"First validation problem: {result.ValidationProblems.FirstOrDefault()?.Message}");
        Assert.That(result.Data.Count, Is.EqualTo(5));
        
        Assert.That(result.Data[0].FirstColumn, Is.EqualTo(1d));
        Assert.That(result.Data[1].FirstColumn, Is.EqualTo(2d));
        Assert.That(result.Data[2].FirstColumn, Is.EqualTo(4d));
        Assert.That(result.Data[3].FirstColumn, Is.EqualTo(5d));
        Assert.That(result.Data[4].FirstColumn, Is.EqualTo(6d));
        
        Assert.That(result.Data[0].SecondColumn, Is.EqualTo("one"));
        Assert.That(result.Data[1].SecondColumn, Is.EqualTo("two"));
        Assert.That(result.Data[2].SecondColumn, Is.EqualTo("four"));
        Assert.That(result.Data[3].SecondColumn, Is.EqualTo("five"));
        Assert.That(result.Data[4].SecondColumn, Is.EqualTo("six"));
    }
    
    [Worksheet(Name = "WithBlankRows")]
    private class WorksheetWithBlankRowsAndAllOptionalProperties
    {
        [Column(Optional = true)] public double? FirstColumn { get; init; }
        [Column(Optional = true)] public string? SecondColumn { get; init; }
    }
    
    [Test]
    public void Given_WorksheetWithBlankRowsAndAllPropertiesOptional_Then_TheBlankRowsResultInAnObjectWithNoPropertiesSet()
    {
        var result = ExcelToObjects.ReadData<WorksheetWithBlankRowsAndAllOptionalProperties>(_testFilename);
        Assert.That(result.IsValid, Is.True, $"First validation problem: {result.ValidationProblems.FirstOrDefault()?.Message}");
        Assert.That(result.Data.Count, Is.EqualTo(7));
        
        Assert.That(result.Data[0].FirstColumn, Is.EqualTo(1d));
        Assert.That(result.Data[1].FirstColumn, Is.EqualTo(2d));
        Assert.That(result.Data[2].FirstColumn, Is.Null);
        Assert.That(result.Data[3].FirstColumn, Is.EqualTo(4d));
        Assert.That(result.Data[4].FirstColumn, Is.Null);
        Assert.That(result.Data[5].FirstColumn, Is.EqualTo(5d));
        Assert.That(result.Data[6].FirstColumn, Is.EqualTo(6d));
        
        Assert.That(result.Data[0].SecondColumn, Is.EqualTo("one"));
        Assert.That(result.Data[1].SecondColumn, Is.EqualTo("two"));
        Assert.That(result.Data[2].SecondColumn, Is.Null);
        Assert.That(result.Data[3].SecondColumn, Is.EqualTo("four"));
        Assert.That(result.Data[4].SecondColumn, Is.Null);
        Assert.That(result.Data[5].SecondColumn, Is.EqualTo("five"));
        Assert.That(result.Data[6].SecondColumn, Is.EqualTo("six"));
    }
    
    [Worksheet(Name = "Sheet1", HasHeadings = true)]
    private class BenchmarkData
    {
        // ReSharper disable UnusedMember.Local
        [Column] public int First { get; set; }
        [Column] public double Second { get; set; }
        [Column] public string Third { get; set; }
        [Column] public string Fourth { get; set; }
        [Column] public DateOnly Fifth { get; set; }
        // ReSharper restore UnusedMember.Local
    }

    [Test]
    public void BenchmarkTest()
    {
        var filename = Path.Combine(Environment.CurrentDirectory, @"TestFiles\Benchmark.xlsx");
        var results = ExcelToObjects.ReadData<BenchmarkData>(filename);
        Assert.That(results.IsValid, Is.True);
    }
}
