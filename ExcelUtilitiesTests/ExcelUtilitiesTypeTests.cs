namespace ExcelUtilities;

public class ExcelUtilitiesTypeTests
{
    private string _testFilename = null!;

    [SetUp]
    public void Setup()
    {
        _testFilename = Path.Combine(Environment.CurrentDirectory, @"TestFiles\FirstTest.xlsx");
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
    private class InvalidDoubleColumnWorksheet
    {
        [Column(Name = "A")] public double DoubleColumn { get; init; }
    }
    
    [Test]
    public void Given_AStringColumnMappedToADoubleProperty_Then_AValidationProblemWillReportAMappingError()
    {
        var result = ExcelToObjects.ReadData<InvalidDoubleColumnWorksheet>(_testFilename);
        Assert.That(result.IsValid, Is.False, "The string data should not map to the double type.");
        Assert.That(result.ValidationProblems[0].Message, Is.EqualTo("The cell TypeTests!A1 has the value 'one' which cannot be interpreted as the data type 'Double'."));
    }

    [Worksheet(Name = "TypeTests")]
    private class IntegerColumnWorksheet
    {
        [Column(Name = "H")] public int IntegerColumn { get; init; }
        [Column(Name = "H")] public int? NullableIntegerColumn { get; init; }
    }
    
    [Test]
    public void Given_ColumnDefinitionOfInteger_Then_TheDataIsMappedCorrectly()
    {
        var result = ExcelToObjects.ReadData<IntegerColumnWorksheet>(_testFilename);
        Assert.That(result.IsValid, Is.True, $"First validation problem: {result.ValidationProblems.FirstOrDefault()?.Message}");
        Assert.That(result.Data[0].IntegerColumn, Is.EqualTo(3));
        Assert.That(result.Data[0].NullableIntegerColumn, Is.EqualTo(3));
    }

    [Worksheet(Name = "TypeTests")]
    private class InvalidIntegerColumnWorksheet
    {
        [Column(Name = "A")] public int IntegerColumn { get; init; }
    }
    
    [Test]
    public void Given_AStringColumnMappedToAIntegerProperty_Then_AValidationProblemWillReportAMappingError()
    {
        var result = ExcelToObjects.ReadData<InvalidIntegerColumnWorksheet>(_testFilename);
        Assert.That(result.IsValid, Is.False, "The string data should not map to the int type.");
        Assert.That(result.ValidationProblems[0].Message, Is.EqualTo("The cell TypeTests!A1 has the value 'one' which cannot be interpreted as the data type 'Int32'."));
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
    private class InvalidDateColumnWorksheet
    {
        [Column(Name = "A")] public DateOnly DateColumn { get; init; }
    }
    
    [Test]
    public void Given_AStringColumnMappedToADateProperty_Then_AValidationProblemWillReportAMappingError()
    {
        var result = ExcelToObjects.ReadData<InvalidDateColumnWorksheet>(_testFilename);
        Assert.That(result.IsValid, Is.False, "The string data should not map to the DateOnly type.");
        Assert.That(result.ValidationProblems[0].Message, Is.EqualTo("The cell TypeTests!A1 has the value 'one' which cannot be interpreted as the data type 'DateOnly'."));
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
    private class InvalidDateTimeColumnWorksheet
    {
        [Column(Name = "A")] public DateTime DateTimeColumn { get; init; }
    }
    
    [Test]
    public void Given_AStringColumnMappedToADateTimeProperty_Then_AValidationProblemWillReportAMappingError()
    {
        var result = ExcelToObjects.ReadData<InvalidDateTimeColumnWorksheet>(_testFilename);
        Assert.That(result.IsValid, Is.False, "The string data should not map to the DateTime type.");
        Assert.That(result.ValidationProblems[0].Message, Is.EqualTo("The cell TypeTests!A1 has the value 'one' which cannot be interpreted as the data type 'DateTime'."));
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
    private class InvalidTimeOnlyColumnWorksheet
    {
        [Column(Name = "A")] public TimeOnly TimeOnlyColumn { get; init; }
    }
    
    [Test]
    public void Given_AStringColumnMappedToATimeOnlyProperty_Then_AValidationProblemWillReportAMappingError()
    {
        var result = ExcelToObjects.ReadData<InvalidTimeOnlyColumnWorksheet>(_testFilename);
        Assert.That(result.IsValid, Is.False, "The string data should not map to the TimeOnly type.");
        Assert.That(result.ValidationProblems[0].Message, Is.EqualTo("The cell TypeTests!A1 has the value 'one' which cannot be interpreted as the data type 'TimeOnly'."));
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
    private class InvalidAccountingColumnWorksheet
    {
        [Column(Name = "A")] public double AccountingColumn { get; init; }
    }
    
    [Test]
    public void Given_AStringColumnMappedToAnAccountingProperty_Then_AValidationProblemWillReportAMappingError()
    {
        var result = ExcelToObjects.ReadData<InvalidAccountingColumnWorksheet>(_testFilename);
        Assert.That(result.IsValid, Is.False, "The string data should not map to the double type.");
        Assert.That(result.ValidationProblems[0].Message, Is.EqualTo("The cell TypeTests!A1 has the value 'one' which cannot be interpreted as the data type 'Double'."));
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

    [Worksheet(Name = "TypeTests")]
    private class InvalidCurrencyColumnWorksheet
    {
        [Column(Name = "A")] public double CurrencyColumn { get; init; }
    }
    
    [Test]
    public void Given_AStringColumnMappedToAnCurrencyProperty_Then_AValidationProblemWillReportAMappingError()
    {
        var result = ExcelToObjects.ReadData<InvalidCurrencyColumnWorksheet>(_testFilename);
        Assert.That(result.IsValid, Is.False, "The string data should not map to the double type.");
        Assert.That(result.ValidationProblems[0].Message, Is.EqualTo("The cell TypeTests!A1 has the value 'one' which cannot be interpreted as the data type 'Double'."));
    }
}