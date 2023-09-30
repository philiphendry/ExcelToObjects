using ClosedXML.Excel;

namespace ExcelToObjects;

public static class ExcelToObjects
{
    public static ConversionResult<T> ReadData<T>(string filename) where T : new()
    {
        var isValid = true;
        var validationProblems = new List<ValidationProblem>();
        var data = new List<T>();

        var worksheetAttribute = typeof(T).GetCustomAttributes(typeof(WorksheetAttribute), false).SingleOrDefault() as WorksheetAttribute;
        var worksheetName = worksheetAttribute is not null && !string.IsNullOrEmpty(worksheetAttribute.Name)
            ? worksheetAttribute.Name
            : typeof(T).Name;

        var workbook = new XLWorkbook(filename);
        if (!workbook.Worksheets.TryGetWorksheet(worksheetName, out var worksheet))
        {
            isValid = false;
            validationProblems.Add(new ValidationProblem($"The worksheet could not be found with the name '{worksheetName}'."));
        }
        else
        {
            var columnProperties = typeof(T).GetProperties().Select((p, i) =>
                new
                {
                    PropertyName = p.Name,
                    PropertyIndex = i,
                    PropertyInfo = p,
                    ColumnDefinition =
                        p.GetCustomAttributes(typeof(ColumnAttribute), false).SingleOrDefault() as ColumnAttribute
                        ?? new ColumnAttribute { Index = i },
                    ColumnIndex = 1
                })
                .ToList();

            var rowCount = worksheet.LastRowUsed().RowNumber();
            var columnCount = worksheet.LastColumnUsed().ColumnNumber();

            var row = 1;
            while (row <= rowCount)
            {
                var dataRow = new T();
                data.Add(dataRow);

                foreach (var columnProperty in columnProperties)
                {
                    var cellValue = worksheet.Cell(row, columnProperty.ColumnIndex);
                    columnProperty.PropertyInfo.SetValue(dataRow, cellValue.GetString());
                }
                
                row++;
            }
        }

        return new ConversionResult<T>(isValid, validationProblems, data);
    }
}

public class ConversionResult<T>
{
    public ConversionResult(bool isValid, List<ValidationProblem> validationProblems, List<T> data)
    {
        IsValid = isValid;
        ValidationProblems = validationProblems ?? throw new ArgumentNullException(nameof(validationProblems));
        Data = data ?? throw new ArgumentNullException(nameof(data));
    }

    public bool IsValid { get; }

    public List<ValidationProblem> ValidationProblems { get; }
    
    public List<T> Data { get; }
}

public class ValidationProblem
{
    public string Message { get; }

    public ValidationProblem(string message)
    {
        Message = message;
    }
}