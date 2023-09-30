using System.Reflection;
using ExcelUtilities;

namespace ExcelToObjectsTests;

public class ColumnIndexesTests
{
    private class TestColumnDefinitions
    {
        [Column(Index = 3)] public string? ColumnWithIndex { get; set; }
        [Column(Name = "BBBB")] public string? ColumnWithNameWrongFormat { get; set; }
        [Column(Name = "BB")] public string? ColumnWithName { get; set; }

        // ReSharper disable once InconsistentNaming
        [Column] public string? DQ { get; set; }

        [Column(Heading = "four")] public string? ColumnWithHeading { get; set; }

        [Column(Heading = "not a match")] public string? ColumnWithNoMatchingHeading { get; set; }
        
        [Column] public string? Three { get; set; }
        
        [Column] public string? MatchedByPropertyIndex { get; set; }
    }

    [Test]
    public void Given_AnAttributeSpecifiesTheIndex_Then_TheProvidedIndexIsReturned()
    {
        var propertyInfo = typeof(TestColumnDefinitions).GetProperty(nameof(TestColumnDefinitions.ColumnWithIndex))!;
        var index = ColumnIndexes.GetColumnIndex(propertyInfo, 0, Array.Empty<string>());
        Assert.That(index, Is.EqualTo(3));
    }

    [Test]
    public void Given_AnAttributeNameIsProvidedButInTheWrongFormat_Then_AnExceptionIsThrown()
    {
        var propertyInfo =
            typeof(TestColumnDefinitions).GetProperty(nameof(TestColumnDefinitions.ColumnWithNameWrongFormat))!;
        var exception = Assert.Throws<InvalidOperationException>(() =>
            ColumnIndexes.GetColumnIndex(propertyInfo, 0, Array.Empty<string>()));
        Assert.That(exception!.Message,
            Is.EqualTo("The property 'ColumnAttribute.ColumnWithNameWrongFormat' has an invalid Name of 'BBBB'."));
    }

    [Test]
    public void Given_AnAttributeNameIsProvidedAndExistsInTheHeadings_Then_TheMatchingColumnIndexIsReturned()
    {
        var propertyInfo = typeof(TestColumnDefinitions).GetProperty(nameof(TestColumnDefinitions.ColumnWithName))!;
        var index = ColumnIndexes.GetColumnIndex(propertyInfo, 0, Array.Empty<string>());
        Assert.That(index, Is.EqualTo(54));
    }

    [Test]
    public void Given_APropertyNamedAsAnExcelColumn_Then_TheMatchingColumnIndexIsReturned()
    {
        var propertyInfo = typeof(TestColumnDefinitions).GetProperty(nameof(TestColumnDefinitions.DQ))!;
        var index = ColumnIndexes.GetColumnIndex(propertyInfo, 0, Array.Empty<string>());
        Assert.That(index, Is.EqualTo(121));
    }

    [Test]
    public void Given_AnAttributeHeadingIsProvidedAndExistsInTheHeadings_Then_TheMatchingColumnIndexIsReturned()
    {
        var propertyInfo = typeof(TestColumnDefinitions).GetProperty(nameof(TestColumnDefinitions.ColumnWithHeading))!;
        var index = ColumnIndexes.GetColumnIndex(propertyInfo, 0, new[] { "one", "two", "three", "four", "five", "six" });
        Assert.That(index, Is.EqualTo(4));
    }

    private class TestFirstHeadingIndex
    {
        [Column(Heading = "one")] public string FirstColumn { get; set; }
    }
    
    [Test]
    public void Given_AnAttributeHeadingIsProvidedMatchingTheFirstHeading_Then_IndexOneIsReturned()
    {
        var propertyInfo = typeof(TestFirstHeadingIndex).GetProperty(nameof(TestFirstHeadingIndex.FirstColumn))!;
        var index = ColumnIndexes.GetColumnIndex(propertyInfo, 0, new[] { "one", "two", "three", "four", "five", "six" });
        Assert.That(index, Is.EqualTo(1));
    }

    [Test]
    public void Given_AnAttributeHeadingIsProvidedThatDoesNotExist_Then_AnExceptionIsThrown()
    {
        var propertyInfo = typeof(TestColumnDefinitions).GetProperty(nameof(TestColumnDefinitions.ColumnWithNoMatchingHeading))!;
        var exception = Assert.Throws<InvalidOperationException>(() => ColumnIndexes.GetColumnIndex(propertyInfo, 0, new[] { "one", "two", "three", "four", "five", "six" }));
        Assert.That(exception!.Message, Is.EqualTo("The property 'ColumnAttribute.Heading' has provided a heading 'not a match' that does not exist in the list of spreadsheet headings."));
    }
    
    [Test]
    public void Given_APropertyNameThatMatchesAHeading_Then_TheMatchingColumnIndexIsReturned()
    {
        var propertyInfo = typeof(TestColumnDefinitions).GetProperty(nameof(TestColumnDefinitions.Three))!;
        var index = ColumnIndexes.GetColumnIndex(propertyInfo, 0, new[] { "one", "two", "three", "four", "five", "six" });
        Assert.That(index, Is.EqualTo(3));
    }
    
    [Test]
    public void Given_AllIdentificationHasFailed_Then_ReturnTheColumnForUsingTheIndexOfThePropertyInTheClass()
    {
        var propertyInfo = typeof(TestColumnDefinitions).GetProperty(nameof(TestColumnDefinitions.MatchedByPropertyIndex))!;
        var propertyIndex = typeof(TestColumnDefinitions)
            .GetProperties()
            .Where(p => Attribute.IsDefined(p, typeof(ColumnAttribute)))
            .OrderBy(p => p.GetCustomAttribute<ColumnAttribute>()!.Order)
            .ToList()
            .FindIndex(pi => pi.Name == nameof(TestColumnDefinitions.MatchedByPropertyIndex));
        var index = ColumnIndexes.GetColumnIndex(propertyInfo, propertyIndex, new[] { "one", "two", "three", "four", "five", "six", "seven", "eight", "nine" });
        Assert.That(index, Is.EqualTo(8));
    }
}