namespace ExcelUtilities;

public class ConversionResult<T>
{
    public ConversionResult(List<ValidationProblem> validationProblems, List<T> data)
    {
        IsValid = !validationProblems.Any();
        ValidationProblems = validationProblems ?? throw new ArgumentNullException(nameof(validationProblems));
        Data = data ?? throw new ArgumentNullException(nameof(data));
    }

    public bool IsValid { get; }

    public List<ValidationProblem> ValidationProblems { get; }
    
    public List<T> Data { get; }
}