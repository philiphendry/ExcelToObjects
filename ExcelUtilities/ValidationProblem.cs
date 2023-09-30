namespace ExcelUtilities;

public class ValidationProblem
{
    public string Message { get; }

    public ValidationProblem(string message)
    {
        Message = message;
    }
}