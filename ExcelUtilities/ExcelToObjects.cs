using Sylvan.Data.Excel;

namespace ExcelUtilities;

public static class ExcelToObjects
{
    public static ConversionResult<T> ReadData<T>(string filename) where T : new()
    {
        using Stream stream = File.Open(filename, FileMode.Open);
        return InternalReadData<T>(stream);
    }
    
    public static ConversionResult<T> ReadData<T>(Stream spreadsheetStream) where T : new()
    {
        return InternalReadData<T>(spreadsheetStream);
    }

    private static ConversionResult<T> InternalReadData<T>(Stream spreadsheetStream) where T : new()
    {
        var worksheetAttribute = typeof(T).GetCustomAttributes(typeof(WorksheetAttribute), false).SingleOrDefault() as WorksheetAttribute
                                 ?? new WorksheetAttribute { Name = typeof(T).Name };

        var schemaProvider = worksheetAttribute.HasHeadings ? ExcelSchema.Default : ExcelSchema.NoHeaders;
        var excelDataReaderOptions = new ExcelDataReaderOptions { Schema = schemaProvider };
        using ExcelDataReader excelDataReader = ExcelDataReader.Create(spreadsheetStream, ExcelWorkbookType.ExcelXml, excelDataReaderOptions);

        if (!excelDataReader.TryOpenWorksheet(worksheetAttribute.Name!))
        {
            return new ConversionResult<T>(
                new List<ValidationProblem>
                    { new($"The worksheet could not be found with the name '{worksheetAttribute.Name}'.") },
                new List<T>());
        }
         
        var worksheetResult = LoadWorksheet<T>(worksheetAttribute, excelDataReader);
        return new ConversionResult<T>(worksheetResult.validationProblems, worksheetResult.data);
    }

    private static (List<T> data, List<ValidationProblem> validationProblems) LoadWorksheet<T>(
        WorksheetAttribute worksheetAttribute, 
        ExcelDataReader worksheet) where T : new()
    {
        var validationProblems = new List<ValidationProblem>();
        var data = new List<T>();

        var worksheetHeadings = GetWorksheetHeadings<T>(worksheetAttribute, worksheet);
        var propertyMappings = GetPropertyMappings<T>(worksheetHeadings);

        int? firstBlankRow = null;
        bool areAllPropertiesOptional = propertyMappings.All(p => p.Optional);
        while (worksheet.Read() && worksheet.RowNumber <= worksheet.RowCount)
        {
            if (worksheet.RowFieldCount == 0)
            {
                if (worksheetAttribute.SkipBlankRows)
                {
                    continue;
                }
                
                // If we've just counted blanks rows to the last row then we can ignore them
                if (worksheet.RowNumber == worksheet.RowCount)
                {
                    break;
                }
                
                // If all properties are optional and we're not skipping then process them
                if (!areAllPropertiesOptional)
                {
                    firstBlankRow ??= worksheet.RowNumber;
                    continue;
                }
            }

            if (firstBlankRow.HasValue)
            {
                var firstRequiredProperty = propertyMappings.FirstOrDefault(p => p.Optional == false);
                    
                // If all fields are optional then the blank row is valid and will be processed
                if (firstRequiredProperty != null)
                {
                    validationProblems.Add(new ValidationProblem(
                        $"The cell {worksheet.WorksheetName}!{Utilities.ExcelColumnOrdinalToName(firstRequiredProperty.ColumnIndex + 1)} has no value but is required."));
                    break;
                }
            }

            var dataRow = new T();
            
            LoadCellData(worksheet, propertyMappings, validationProblems, dataRow);

            data.Add(dataRow);
        }

        return (data, validationProblems);
    }

    private static void LoadCellData<T>(
        ExcelDataReader worksheet, List<PropertyMapping> propertyMappings, 
        List<ValidationProblem> validationProblems,
        T dataRow) where T : new()
    {
        foreach (var propertyMapping in propertyMappings)
        {
            if (worksheet.GetValue(propertyMapping.ColumnIndex) == DBNull.Value)
            {
                if (propertyMapping.Optional)
                {
                    continue;
                }

                validationProblems.Add(new ValidationProblem($"The cell {worksheet.WorksheetName}!{Utilities.ExcelColumnOrdinalToName(propertyMapping.ColumnIndex + 1)} has no value but is required."));
                break;
            }

            SetProperty(dataRow, worksheet, propertyMapping);
        }
    }

    private static string[] GetWorksheetHeadings<T>(WorksheetAttribute worksheetAttribute, ExcelDataReader worksheet)
        where T : new()
    {
        string[] worksheetHeadings;
        if (worksheetAttribute.HasHeadings)
        {
            while (worksheet.RowNumber != worksheetAttribute.HeadingsOnRow && worksheet.Read())
            {
            }
            
            worksheetHeadings = new string[worksheet.RowFieldCount];
            for (var columnIndex = 0; columnIndex < worksheet.RowFieldCount; columnIndex++)
            {
                worksheetHeadings[columnIndex] = worksheet.GetString(columnIndex);
            }
        }
        else
        {
            worksheetHeadings = Array.Empty<string>();
        }

        return worksheetHeadings;
    }

    private static void SetProperty<T>(T dataRow, ExcelDataReader excelDataReader, PropertyMapping propertyMapping)
        where T : new()
    {
        var propertyInfo = propertyMapping.PropertyInfo;
        var columnIndex = propertyMapping.ColumnIndex;
        if (propertyInfo.PropertyType == typeof(double))
        {
            propertyInfo.SetValue(dataRow, excelDataReader.GetDouble(columnIndex));
        }
        else if (propertyInfo.PropertyType == typeof(double?))
        {
            propertyInfo.SetValue(dataRow, excelDataReader.GetDouble(columnIndex));
        }
        else if (propertyInfo.PropertyType == typeof(int) || propertyInfo.PropertyType == typeof(int?))
        {
            propertyInfo.SetValue(dataRow, excelDataReader.GetInt32(columnIndex));
        }
        else if (propertyInfo.PropertyType == typeof(DateOnly) || propertyInfo.PropertyType == typeof(DateOnly?))
        {
            propertyInfo.SetValue(dataRow, DateOnly.FromDateTime(excelDataReader.GetDateTime(columnIndex)));
        }
        else if (propertyInfo.PropertyType == typeof(DateTime) || propertyInfo.PropertyType == typeof(DateTime?))
        {
            propertyInfo.SetValue(dataRow, excelDataReader.GetDateTime(columnIndex));
        }
        else if (propertyInfo.PropertyType == typeof(TimeOnly) || propertyInfo.PropertyType == typeof(TimeOnly?))
        {
            propertyInfo.SetValue(dataRow, TimeOnly.FromTimeSpan(excelDataReader.GetTimeSpan(columnIndex)));
        }
        else if (propertyInfo.PropertyType == typeof(string))
        {
            propertyInfo.SetValue(dataRow, excelDataReader.GetString(columnIndex));
        }
        else
        {
            throw new InvalidOperationException(
                $"The property '{typeof(T).Name}.{propertyInfo.Name}' is declared as '{propertyInfo.PropertyType.Name}' which is not supported.");
        }
    }

    private static List<PropertyMapping> GetPropertyMappings<T>(string[] worksheetHeadings) where T : new()
    {
        return typeof(T)
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
                new PropertyMapping(
                    c.PropertyInfo,
                    c.PropertyInfo.Name,
                    propertyIndex,
                    c.ColumnAttribute!.Optional,
                    ColumnIndexes.GetColumnIndex(c.ColumnAttribute!,
                        c.PropertyInfo.Name, 
                        propertyIndex, 
                        worksheetHeadings)))
            // Filter out the columns that are optional and don't exist
            .Where(p => p.ColumnIndex != -1)
            .ToList();
    }
}