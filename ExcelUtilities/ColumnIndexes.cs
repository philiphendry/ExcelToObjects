using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
[assembly:InternalsVisibleTo("ExcelUtilitiesTests")]

namespace ExcelUtilities;

internal static class ColumnIndexes
{
    private static readonly Regex ExcelColumnNameRegex = new Regex("^[A-Z]{1,3}$", RegexOptions.Compiled);

    internal static int GetColumnIndex(PropertyInfo propertyInfo, int propertyIndex, string[] headings)
    {
        var columnAttribute =
            propertyInfo.GetCustomAttributes(typeof(ColumnAttribute), false).SingleOrDefault() as ColumnAttribute
            ?? new ColumnAttribute { Index = propertyIndex };

        if (columnAttribute.Index > 0)
        {
            return columnAttribute.Index;
        }

        if (!string.IsNullOrEmpty(columnAttribute.Name))
        {
            if (!ExcelColumnNameRegex.IsMatch(columnAttribute.Name))
            {
                throw new InvalidOperationException($"The property '{nameof(ColumnAttribute)}.{propertyInfo.Name}' has an invalid Name of '{columnAttribute.Name}'.");
            }

            return Utilities.ExcelColumnNameToIndex(columnAttribute.Name);
        }

        if (ExcelColumnNameRegex.IsMatch(propertyInfo.Name))
        {
            return Utilities.ExcelColumnNameToIndex(propertyInfo.Name);
        }

        if (headings.Any() && !string.IsNullOrEmpty(columnAttribute.Heading))
        {
            var headingIndex = Array.FindIndex(headings, h => h.Equals(columnAttribute.Heading, StringComparison.InvariantCultureIgnoreCase));
            if (headingIndex >= 0)
            {
                return headingIndex + 1;
            }

            if (columnAttribute.Optional)
            {
                return -1;
            }
            
            throw new InvalidOperationException($"The property '{nameof(ColumnAttribute)}.{nameof(ColumnAttribute.Heading)}' has provided a heading '{columnAttribute.Heading}' that does not exist in the list of spreadsheet headings.");
        }
        
        var propertyNameIndex = Array.FindIndex(headings, h => h.Equals(propertyInfo.Name, StringComparison.InvariantCultureIgnoreCase));
        if (propertyNameIndex >= 0)
        {
            return propertyNameIndex + 1;
        }

        return propertyIndex + 1;
    }
}