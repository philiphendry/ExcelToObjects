namespace ExcelUtilities;

internal static class Utilities
{
    internal static int ExcelColumnNameToOrdinal(string columnName)
    {
        var index = 0;
        foreach (var letter in columnName)
        {
            index *= 26;
            index += letter - 'A' + 1;
        }

        return index;
    }
    
    internal static string ExcelColumnOrdinalToName(int columnOrdinal)
    {
        var dividend = columnOrdinal;
        var columnName = string.Empty;
	
        while (dividend > 0)
        {
            var modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }
        return columnName;
    }
}