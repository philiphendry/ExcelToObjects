using System.Runtime.CompilerServices;

namespace ExcelUtilities;

/// <summary>
/// Provides information to map data from the spreadsheet to the attributed
/// property.
///
/// The column is identified by one of <see cref="Heading"/>, <see cref="Name"/>,
/// or <see cref="Index"/>. More than one can be provided and the one used is
/// determined by the following rules
///<list type="number">
///     <item>
///         If <see cref="Index"/> is provided this is used.
///     </item>
///     <item>
///         If <see cref="Name"/> is provided or the property name is a column
///         name (such that it is named with 1, 2, or 3 uppercase letters) then
///         this is used. Column names are, for example, 'A', 'B', or 'AA' etc.
///     </item>
///     <item>
///         If <see cref="Heading"/> is provided and <see cref="WorksheetAttribute.HasHeadings"/>
///         is enabled then the column containing the provided heading is used.
///         An exception will be thrown if the heading does not exist.
///     </item>
///     <item>
///         Finally, the index position of the property in the class will be used as
///         the column index.
///     </item>
///</list>
/// Cell data from the spreadsheet will be coerced to the type the property is
/// declared as. If the conversion cannot be made a validation error will be reported.
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public class ColumnAttribute : Attribute
{
    public ColumnAttribute([CallerLineNumber]int order = 0)
    {
        Order = order;
    }
    
    /// <summary>
    /// An internal property used to determine the order of Columns by there position in the class.
    /// </summary>
    internal int Order { get; }

    /// <summary>
    /// The heading this column is identified by.
    /// <see cref="WorksheetAttribute.HasHeadings"/> must be set for
    /// this property to have effect.
    /// </summary>
    public string? Heading { get; init; }

    /// <summary>
    /// The column as identified by the built-in names. For example, the
    /// first column is 'A' followed by 'B' and the 27th column 'AA'.
    /// </summary>
    public string? Name { get; init; }

    /// <summary>
    /// The column identified by its index. This is one-based so column
    /// 'A' is at index 1.
    /// </summary>
    public int Index { get; init; }

    /// <summary>
    /// By default all columns are required but can be made optional
    /// by set <see cref="Optional"/> to true;
    /// </summary>
    public bool Optional { get; init; }
}