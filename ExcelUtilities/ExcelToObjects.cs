using System.Diagnostics.CodeAnalysis;
using Sylvan.Data.Excel;

namespace ExcelUtilities;

public class ExcelToObjects
{
    public delegate void SetPropertyHandler<in T>(T dataRow, ExcelDataReader excelDataReader, PropertyMapping propertyMapping);
    
    private readonly Dictionary<Type, SetPropertyHandler<object>> _propertyHandlers;

    public ExcelToObjects()
    {
        _propertyHandlers = new Dictionary<Type, SetPropertyHandler<object>>
        {
            { typeof(double), SetPropertyDouble },
            { typeof(double?), SetPropertyDouble },
            { typeof(int), SetPropertyInt },
            { typeof(int?), SetPropertyInt },
            { typeof(DateOnly), SetPropertyDate },
            { typeof(DateOnly?), SetPropertyDate },
            { typeof(DateTime), SetPropertyDateTime },
            { typeof(DateTime?), SetPropertyDateTime },
            { typeof(TimeOnly), SetPropertyTime },
            { typeof(TimeOnly?), SetPropertyTime },
            { typeof(string), SetPropertyString }
        };
    }
    
    /// <summary>
    /// By default the following types are supported:
    /// <list type="bullet">
    ///     <item><description>double</description></item>
    ///     <item><description>int</description></item>
    ///     <item><description>DateOnly</description></item>
    ///     <item><description>DateTime</description></item>
    ///     <item><description>TimeOnly</description></item>
    ///     <item><description>string</description></item>
    /// </list>
    /// Along with their nullable counterparts. If you need to support additional types then you can add a handler for them.
    /// 
    /// </summary>
    /// <param name="type">The new type to add support for.</param>
    /// <param name="handler">A function to perform conversion and setting the property on the target object.</param>
    /// <exception cref="ArgumentNullException"></exception>
    /// <exception cref="ArgumentException"></exception>
    public void AddPropertyHandler(Type type, SetPropertyHandler<object> handler)
    {
        if (type == null) throw new ArgumentNullException(nameof(type));
        if (handler == null) throw new ArgumentNullException(nameof(handler));
        if (!_propertyHandlers.TryAdd(type, handler))
        {
            throw new ArgumentException($"A handler for '{type.Name}' already exists.", nameof(type));
        }
    }

    /// <summary>
    /// Reads the content of the Excel spreadsheet given by the filename and converts it to a list of objects of type T.
    /// If the file does not exist then a FileNotFoundException will be thrown. If is not valid and cannot be
    /// processed according to the attributes then a ConversionResult will be returned with the validation problems.
    /// </summary>
    /// <param name="filename"></param>
    /// <typeparam name="T"></typeparam>
    /// <returns></returns>
    /// <exception cref="ArgumentNullException"></exception>
    /// <exception cref="FileNotFoundException"></exception>
    public ConversionResult<T> ReadData<T>(string filename) where T : new()
    {
        if (filename is null) throw new ArgumentNullException(nameof(filename));

        if (!File.Exists(filename))
        {
            throw new FileNotFoundException("The file could not be found.", filename);
        }
        
        using Stream stream = File.Open(filename, FileMode.Open);
        return InternalReadData<T>(stream);
    }
    
    /// <summary>
    /// Reads the content of the Excel spreadsheet given by the stream and converts it to a list of objects of type T.
    /// </summary>
    /// <param name="spreadsheetStream"></param>
    /// <typeparam name="T"></typeparam>
    /// <returns></returns>
    public ConversionResult<T> ReadData<T>(Stream spreadsheetStream) where T : new()
    {
        return InternalReadData<T>(spreadsheetStream);
    }

    private ConversionResult<T> InternalReadData<T>(Stream spreadsheetStream) where T : new()
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

    private (List<T> data, List<ValidationProblem> validationProblems) LoadWorksheet<T>(
        WorksheetAttribute worksheetAttribute, 
        ExcelDataReader worksheet) where T : new()
    {
        var validationProblems = new List<ValidationProblem>();
        var data = new List<T>();

        var worksheetHeadings = GetWorksheetHeadings(worksheetAttribute, worksheet);
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

    private void LoadCellData<T>(
        ExcelDataReader worksheet, List<PropertyMapping> propertyMappings, 
        List<ValidationProblem> validationProblems,
        [DisallowNull] T dataRow) where T : new()
    {
        if (dataRow == null) throw new ArgumentNullException(nameof(dataRow));
        
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

    private static string[] GetWorksheetHeadings(WorksheetAttribute worksheetAttribute, ExcelDataReader worksheet)
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

    private void SetProperty<T>([DisallowNull] T dataRow, ExcelDataReader excelDataReader, PropertyMapping propertyMapping)
        where T : new()
    {
        if (dataRow == null) throw new ArgumentNullException(nameof(dataRow));
        
        if (_propertyHandlers.TryGetValue(propertyMapping.PropertyInfo.PropertyType, out var handler))
        {
            handler(dataRow, excelDataReader, propertyMapping);
        }
        else
        {
            throw new InvalidOperationException($"The property '{typeof(T).Name}.{propertyMapping.PropertyInfo.Name}' is declared as '{propertyMapping.PropertyInfo.PropertyType.Name}' which is not supported.");
        }
    }

    private static void SetPropertyString<T>(T dataRow, ExcelDataReader excelDataReader, PropertyMapping propertyMapping)
        where T : new()
    {
        propertyMapping.PropertyInfo.SetValue(dataRow, excelDataReader.GetString(propertyMapping.ColumnIndex));
    }

    private static void SetPropertyTime<T>(T dataRow, ExcelDataReader excelDataReader, PropertyMapping propertyMapping)
        where T : new()
    {
        propertyMapping.PropertyInfo.SetValue(dataRow, TimeOnly.FromTimeSpan(excelDataReader.GetTimeSpan(propertyMapping.ColumnIndex)));
    }

    private static void SetPropertyDateTime<T>(T dataRow, ExcelDataReader excelDataReader, PropertyMapping propertyMapping)
        where T : new()
    {
        propertyMapping.PropertyInfo.SetValue(dataRow, excelDataReader.GetDateTime(propertyMapping.ColumnIndex));
    }

    private static void SetPropertyDate<T>(T dataRow, ExcelDataReader excelDataReader, PropertyMapping propertyMapping)
        where T : new()
    {
        propertyMapping.PropertyInfo.SetValue(dataRow, DateOnly.FromDateTime(excelDataReader.GetDateTime(propertyMapping.ColumnIndex)));
    }

    private static void SetPropertyInt<T>(T dataRow, ExcelDataReader excelDataReader, PropertyMapping propertyMapping)
        where T : new()
    {
        propertyMapping.PropertyInfo.SetValue(dataRow, excelDataReader.GetInt32(propertyMapping.ColumnIndex));
    }

    private static void SetPropertyDouble<T>(T dataRow, ExcelDataReader excelDataReader, PropertyMapping propertyMapping)
        where T : new()
    {
        propertyMapping.PropertyInfo.SetValue(dataRow, excelDataReader.GetDouble(propertyMapping.ColumnIndex));
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