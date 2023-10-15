using System.Reflection;

namespace ExcelUtilities;

/// <summary>
/// Represents the data read from <see cref="ColumnAttribute"/> decorating the target class
/// but enhanced with details read from the spreadsheet.
/// </summary>
/// <param name="PropertyInfo"></param>
/// <param name="PropertyName"></param>
/// <param name="PropertyIndex"></param>
/// <param name="Optional"></param>
/// <param name="ColumnIndex"></param>
public record PropertyMapping(PropertyInfo PropertyInfo, string PropertyName, int PropertyIndex, bool Optional, int ColumnIndex)
{
    public override string ToString()
    {
        return $"{{ PropertyInfo = {PropertyInfo}, PropertyName = {PropertyName}, PropertyIndex = {PropertyIndex}, Optional = {Optional}, ColumnIndex = {ColumnIndex} }}";
    }
}