namespace ExcelUtilities;

/// <summary>
/// Represents the result of a conversion operation from an Excel file to a list of objects of type T.
/// If the conversion is successful, <see cref="IsValid"/> will be true and <see cref="Data"/> will contain the
/// result of the conversion. If unsuccessful, <see cref="IsValid"/> will be false and <see cref="ValidationProblems"/>
/// will contain the validation errors.
/// </summary>
/// <typeparam name="T">
/// The type to convert to and which has been decorated with attributes indicating how to map the data from the spreadsheet.
/// </typeparam>
public class ConversionResult<T>
{
    internal ConversionResult(List<ValidationProblem> validationProblems, List<T> data)
    {
        IsValid = !validationProblems.Any();
        ValidationProblems = validationProblems ?? throw new ArgumentNullException(nameof(validationProblems));
        Data = data ?? throw new ArgumentNullException(nameof(data));
    }

    /// <summary>
    /// If the conversion was successful the value will be true and <see cref="Data"/> will contain the result.
    /// If unsuccessful, the value will be false and <see cref="ValidationProblems"/> will contain the validation errors.
    /// </summary>
    public bool IsValid { get; }

    /// <summary>
    /// The list of validation problems that occurred during the conversion.
    /// </summary>
    public List<ValidationProblem> ValidationProblems { get; }
    
    /// <summary>
    /// The converted data if the conversion was successful.
    /// </summary>
    public List<T> Data { get; }
}