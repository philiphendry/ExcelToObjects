namespace ExcelUtilities;

/// <summary>
/// Provides information to identify a worksheet and map data from
/// it to new instances of the attributed class.
/// </summary>
[AttributeUsage(AttributeTargets.Class)]
public class WorksheetAttribute : Attribute
{
    /// <summary>
    /// The name of the worksheet data will be loaded from.
    /// If not provided the name of the class will be used.
    /// </summary>
    public string? Name { get; init; } = null;

    /// <summary>
    /// If no mapped properties are set from the row in the spreadsheet then, be
    /// default, a validation problem will be raised along with any problems resulting
    /// from required properties that have no value. Setting <see cref="SkipBlankRows"/>
    /// to <value>true</value> will skip these rows without reporting any issues.
    /// </summary>
    public bool SkipBlankRows { get; init; } = false;

    /// <summary>
    /// If the worksheet contains a row with headings then set <see cref="HasHeadings"/>
    /// to <value>true</value>. The default is <value>false</value>. 
    /// </summary>
    public bool HasHeadings { get; init; } = false;

    /// <summary>
    /// If <see cref="HasHeadings"/> is <value>true</value> then they are assumed
    /// to be on the first row (index 1.) Specify the one-based row number where the
    /// headings are to override this.
    /// </summary>
    public int HeadingsOnRow { get; init; } = 1;
}