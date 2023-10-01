using System.Reflection;
using ClosedXML.Excel;

namespace ExcelUtilities;

public static class ExcelToObjects
{
    public static ConversionResult<T> ReadData<T>(string filename) where T : new()
    {
        var worksheetAttribute = typeof(T).GetCustomAttributes(typeof(WorksheetAttribute), false).SingleOrDefault() as WorksheetAttribute
                                 ?? new WorksheetAttribute { Name = typeof(T).Name };
        
        var workbook = new XLWorkbook(filename);
        if (!workbook.Worksheets.TryGetWorksheet(worksheetAttribute.Name, out var worksheet))
        {
            return new ConversionResult<T>(
                new List<ValidationProblem>
                    { new($"The worksheet could not be found with the name '{worksheetAttribute.Name}'.") },
                new List<T>());
        }
        else
        {
            var worksheetResult = ReadWorksheet<T>(worksheetAttribute, worksheet);
            return new ConversionResult<T>(worksheetResult.validationProblems, worksheetResult.data);
        }
    }

    private static (List<T> data, List<ValidationProblem> validationProblems) ReadWorksheet<T>(
        WorksheetAttribute worksheetAttribute, 
        IXLWorksheet worksheet) where T : new()
    {
        var validationProblems = new List<ValidationProblem>();
        var data = new List<T>();

        var worksheetHeadings = worksheetAttribute.HasHeadings
            ? worksheet.Row(worksheetAttribute.HeadingsOnRow).Cells().Select(c => c.Value.ToString()).ToArray()
            : Array.Empty<string>();

        var columnProperties = typeof(T)
            .GetProperties()
            .Select(propertyInfo =>
                new
                {
                    PropertyInfo = propertyInfo,
                    ColumnAttribute =
                        propertyInfo.GetCustomAttributes(typeof(ColumnAttribute), false).SingleOrDefault() as
                            ColumnAttribute
                })
            .Where(c => c.ColumnAttribute != null)
            // Sort the properties by their position in the class so the propertyIndex can be used.
            .OrderBy(c => c.ColumnAttribute!.Order)
            .Select((c, propertyIndex) =>
                new
                {
                    c.PropertyInfo,
                    PropertyName = c.PropertyInfo.Name,
                    PropertyIndex = propertyIndex,
                    c.ColumnAttribute!.Optional,
                    ColumnIndex = ColumnIndexes.GetColumnIndex(c.ColumnAttribute!,
                        c.PropertyInfo.Name, propertyIndex, worksheetHeadings)
                })
            // Filter out the columns that are optional and don't exist
            .Where(p => p.ColumnIndex != -1)
            .ToList();

        var firstRowIndex = worksheetAttribute.HasHeadings ? worksheetAttribute.HeadingsOnRow + 1 : 1;
        var lastRowIndex = worksheet.LastRowUsed(XLCellsUsedOptions.Contents)?.RowNumber() ?? 0;

        // Ignore any trailing blanks rows
        while (lastRowIndex > 0 && worksheet.Row(lastRowIndex).Cells(usedCellsOnly: true)
                   .All(c => string.IsNullOrEmpty(c.GetString().Trim())))
        {
            lastRowIndex--;
        }

        foreach (var row in worksheet.Rows(firstRowIndex, lastRowIndex))
        {
            if (row.Cells(usedCellsOnly: true).All(c => c.DataType == XLDataType.Blank))
            {
                if (worksheetAttribute.SkipBlankRows)
                {
                    continue;
                }

                var firstRequiredProperty = columnProperties.FirstOrDefault(p => p.Optional == false);
                if (firstRequiredProperty != null)
                {
                    validationProblems.Add(new ValidationProblem(
                        $"The cell {worksheet}!{row.Cell(firstRequiredProperty.ColumnIndex)} has no value but is required."));
                    break;
                }
            }

            var dataRow = new T();
            foreach (var columnProperty in columnProperties)
            {
                IXLCell cellValue = null!;
                try
                {
                    cellValue = row.Cell(columnProperty.ColumnIndex);

                    if (cellValue.DataType == XLDataType.Blank)
                    {
                        if (columnProperty.Optional)
                        {
                            continue;
                        }

                        validationProblems.Add(
                            new ValidationProblem($"The cell {worksheet}!{cellValue} has no value but is required."));
                        break;
                    }

                    SetProperty(columnProperty.PropertyInfo, dataRow, cellValue);
                }
                catch (Exception e)
                {
                    throw new InvalidCastException(
                        $"The conversion of the value {worksheet}!{cellValue} with type '{cellValue.DataType}' to property type '{columnProperty.PropertyInfo.PropertyType.Name}' failed.",
                        e);
                }
            }

            if (validationProblems.Count > 0)
            {
                break;
            }

            data.Add(dataRow);
        }

        return (data, validationProblems);
    }

    private static void SetProperty<T>(PropertyInfo propertyInfo, T dataRow, IXLCell cellValue)
        where T : new()
    {
        if (propertyInfo.PropertyType == typeof(double) || propertyInfo.PropertyType == typeof(double?))
        {
            propertyInfo.SetValue(dataRow, cellValue.GetValue<double>());
        }
        else if (propertyInfo.PropertyType == typeof(DateOnly) || propertyInfo.PropertyType == typeof(DateOnly?))
        {
            propertyInfo.SetValue(dataRow, DateOnly.FromDateTime(cellValue.GetDateTime()));
        }
        else if (propertyInfo.PropertyType == typeof(DateTime) || propertyInfo.PropertyType == typeof(DateTime?))
        {
            propertyInfo.SetValue(dataRow, cellValue.GetValue<DateTime>());
        }
        else if (propertyInfo.PropertyType == typeof(TimeOnly) || propertyInfo.PropertyType == typeof(TimeOnly?))
        {
            propertyInfo.SetValue(dataRow, TimeOnly.FromTimeSpan(cellValue.GetValue<TimeSpan>()));
        }
        else if (propertyInfo.PropertyType == typeof(string))
        {
            propertyInfo.SetValue(dataRow, cellValue.GetString());
        }
        else
        {
            throw new InvalidOperationException(
                $"The property '{typeof(T).Name}.{propertyInfo.Name}' is declared as '{propertyInfo.PropertyType.Name}' which is not supported.");
        }
    }
}