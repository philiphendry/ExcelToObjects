namespace ExcelUtilities;

/// <summary>
/// Describes a validation problem that occurred during the conversion of an Excel file to a list of objects.
/// </summary>
public class ValidationProblem
{
    public string Message { get; }

    public string? WorksheetName { get; }

    public string? CellAddress { get; }

    public ValidationProblem(string message)
    {
        Message = message;
    }

    internal ValidationProblem(string message, string? worksheetName, string? cellAddress)
    {
        Message = message;
        WorksheetName = worksheetName;
        CellAddress = cellAddress;
    }
}