namespace ExcelToObjects;

/// <summary>
/// Provides information to map data from the spreadsheet to the attributed
/// property.
///
/// The column is identified by one of <see cref="Heading"/>, <see cref="Name"/>,
/// or <see cref="Index"/>. At most one can be specified and an exception will be
/// thrown if more than one is provided. If none are provided then the property
/// name will be:
///<list type="number">
///   <item>
///     If the property name is a column name such as 'A', 'B', or 'AA' etc. then
///     the column will identified in tha way,
///   </item> 
///   <item>
///     If <see cref="WorksheetAttribute.HasHeadings"/> is true then an attempt
///     to find the property name among the headings is attempted first,
///   </item>
///   <item>
///     Finally, the index position of the property in the class will be used as
///     the column index.
///   </item>
///</list>
/// Cell data from the spreadsheet will be coerced to the type the property is
/// declared as. If the conversion cannot be made a validation error will be reported.
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public class ColumnAttribute : Attribute
{
    /// <summary>
    /// The heading this column is identified by.
    /// <see cref="WorksheetAttribute.HasHeadings"/> must be set for
    /// this property to have effect.
    /// </summary>
    public string? Heading { get; init; } = null;

    /// <summary>
    /// The column as identified by the built-in names. For example, the
    /// first column is 'A' followed by 'B' and the 27th column 'AA'.
    /// </summary>
    public string? Name { get; init; } = null;

    /// <summary>
    /// The column identified by its index. This is one-based so column
    /// 'A' is at index 1.
    /// </summary>
    public int Index { get; init; } = -1;

    /// <summary>
    /// If the column is marked as required by cannot be found a validation
    /// error will be reported. The default is <value>false</value>.
    /// </summary>
    public bool Required { get; init; } = false;
}