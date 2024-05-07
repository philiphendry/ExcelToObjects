using System.Diagnostics;
using System.Reflection;

namespace ExcelUtilities;

/// <summary>
/// Represents the data read from <see cref="ColumnAttribute"/> decorating the target class
/// but enhanced with details read from the spreadsheet.
/// </summary>
/// <param name="PropertyInfo">
/// The PropertyInfo object identifying the property on the target type to which the Excel spreadsheet with column index
/// given by <see cref="ColumnIndex"/> is mapped.
/// </param>
/// <param name="Optional">
/// Set if the <see cref="ColumnAttribute.Optional"/> property was set in the column mapping and indicates that a cell
/// can be empty without causing a validation error.
/// </param>
/// <param name="ColumnIndex">
/// The index of the column in the spreadsheet that the property is mapped to. The mapping is determined by the
/// <see cref="ColumnAttribute.Heading"/>, <see cref="ColumnAttribute.Name"/>, or the <see cref="ColumnAttribute.Index"/>.
/// </param>
[DebuggerDisplay("PropertyInfo={PropertyInfo}, Optional={Optional}, ColumnIndex={ColumnIndex}")]
public record PropertyMapping(PropertyInfo PropertyInfo, bool Optional, int ColumnIndex);