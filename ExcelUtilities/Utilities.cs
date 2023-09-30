namespace ExcelUtilities;

internal static class Utilities
{
    internal static int ExcelColumnNameToIndex(string columnName)
    {
        var index = 0;
        foreach (var letter in columnName)
        {
            index *= 26;
            index += letter - 'A' + 1;
        }

        return index;
    }
}