using ExcelUtilities;

namespace ExcelToObjectsTests;

public class ColumnIndexesTests
{
    [Test]
    public void Given_AnAttributeSpecifiesTheIndex_Then_TheProvidedIndexIsReturned()
    {
        var index = ColumnIndexes.GetColumnIndex(new ColumnAttribute() { Index = 3 }, "ColumnWithIndex", 0, Array.Empty<string>());
        Assert.That(index, Is.EqualTo(3));
    }

    [Test]
    public void Given_AnAttributeNameIsProvidedButInTheWrongFormat_Then_AnExceptionIsThrown()
    {
        var exception = Assert.Throws<InvalidOperationException>(() => ColumnIndexes.GetColumnIndex(new ColumnAttribute() {Name = "BBBB"}, "ColumnWithNameWrongFormat", 0, Array.Empty<string>()));
        Assert.That(exception!.Message, Is.EqualTo("The property 'ColumnAttribute.ColumnWithNameWrongFormat' has an invalid Name of 'BBBB'."));
    }

    [Test]
    public void Given_AnAttributeNameIsProvidedAndExistsInTheHeadings_Then_TheMatchingColumnIndexIsReturned()
    {
        var index = ColumnIndexes.GetColumnIndex(new ColumnAttribute() { Name = "BB"}, "ColumnWithName", 0, Array.Empty<string>());
        Assert.That(index, Is.EqualTo(54));
    }

    [Test]
    public void Given_APropertyNamedAsAnExcelColumn_Then_TheMatchingColumnIndexIsReturned()
    {
        var index = ColumnIndexes.GetColumnIndex(new ColumnAttribute(), "DQ", 0, Array.Empty<string>());
        Assert.That(index, Is.EqualTo(121));
    }

    [Test]
    public void Given_AnAttributeHeadingIsProvidedAndExistsInTheHeadings_Then_TheMatchingColumnIndexIsReturned()
    {
        var index = ColumnIndexes.GetColumnIndex(new ColumnAttribute() { Heading = "Four"}, "ColumnWithHeading", 0, new[] { "one", "two", "three", "four", "five", "six" });
        Assert.That(index, Is.EqualTo(4));
    }

    [Test]
    public void Given_AnAttributeHeadingIsProvidedMatchingTheFirstHeading_Then_IndexOneIsReturned()
    {
        var index = ColumnIndexes.GetColumnIndex(new ColumnAttribute() { Heading = "one"}, "FirstColumn", 0, new[] { "one", "two", "three", "four", "five", "six" });
        Assert.That(index, Is.EqualTo(1));
    }

    [Test]
    public void Given_AnAttributeHeadingIsProvidedThatDoesNotExist_Then_AnExceptionIsThrown()
    {
        var exception = Assert.Throws<InvalidOperationException>(() => ColumnIndexes.GetColumnIndex(new ColumnAttribute() { Heading = "not a match"}, "ColumnWithNoMatchingHeading", 0, new[] { "one", "two", "three", "four", "five", "six" }));
        Assert.That(exception!.Message, Is.EqualTo("The property 'ColumnAttribute.Heading' has provided a heading 'not a match' that does not exist in the list of spreadsheet headings."));
    }
    
    [Test]
    public void Given_APropertyNameThatMatchesAHeading_Then_TheMatchingColumnIndexIsReturned()
    {
        var index = ColumnIndexes.GetColumnIndex(new ColumnAttribute(), "Three", 0, new[] { "one", "two", "three", "four", "five", "six" });
        Assert.That(index, Is.EqualTo(3));
    }
    
    [Test]
    public void Given_AllIdentificationHasFailed_Then_ReturnTheColumnForUsingTheIndexOfThePropertyInTheClass()
    {
        // Property indexes are zero-based whilst the columns are one-based
        var index = ColumnIndexes.GetColumnIndex(new ColumnAttribute(), "MatchedByPropertyIndex", 8, new[] { "one", "two", "three", "four", "five", "six", "seven", "eight", "nine" });
        Assert.That(index, Is.EqualTo(9));
    }
}