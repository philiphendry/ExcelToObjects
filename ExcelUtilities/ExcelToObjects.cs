using ClosedXML.Excel;

namespace ExcelUtilities;

public static class ExcelToObjects
{
    public static ConversionResult<T> ReadData<T>(string filename) where T : new()
    {
        var validationProblems = new List<ValidationProblem>();
        var data = new List<T>();

        var worksheetAttribute = typeof(T).GetCustomAttributes(typeof(WorksheetAttribute), false).SingleOrDefault() as WorksheetAttribute
            ?? new WorksheetAttribute { Name = typeof(T).Name };
        
        var workbook = new XLWorkbook(filename);
        if (!workbook.Worksheets.TryGetWorksheet(worksheetAttribute.Name, out var worksheet))
        {
            validationProblems.Add(new ValidationProblem($"The worksheet could not be found with the name '{worksheetAttribute.Name}'."));
        }
        else
        {
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
                .Select((c, propertyIndex) =>
                    new
                    {
                        PropertyInfo = c.PropertyInfo,
                        PropertyName = c.PropertyInfo.Name,
                        PropertyIndex = propertyIndex,
                        PropertyType = c.PropertyInfo.PropertyType,
                        Optional = c.ColumnAttribute!.Optional,
                        ColumnIndex = ColumnIndexes.GetColumnIndex(c.ColumnAttribute!,
                            c.PropertyInfo.Name, propertyIndex, worksheetHeadings)
                    })
                // Filter out the columns that are optional and don't exist
                .Where(p => p.ColumnIndex != -1)
                .ToList();
            
            var rowCount = worksheet.LastRowUsed(XLCellsUsedOptions.Contents)?.RowNumber() ?? 0;
            var rowIndex = worksheetAttribute.HasHeadings ? worksheetAttribute.HeadingsOnRow + 1 : 1;
            while (rowIndex <= rowCount)
            {
                if (worksheetAttribute.SkipBlankRows 
                    && columnProperties.All(cp =>
                        worksheet.Cell(rowIndex, cp.ColumnIndex).DataType == XLDataType.Blank))
                {
                    rowIndex++;
                    continue;
                }

                var dataRow = new T();
                foreach (var columnProperty in columnProperties)
                {
                    IXLCell cellValue = null!;
                    try
                    {
                        cellValue = worksheet.Cell(rowIndex, columnProperty.ColumnIndex);

                        if (cellValue.DataType == XLDataType.Blank)
                        {
                            if (columnProperty.Optional)
                            {
                                continue;
                            }
                            
                            validationProblems.Add(new ValidationProblem($"The cell {worksheet}!{cellValue} has no value but is required"));
                            break;
                        }
                        
                        if (columnProperty.PropertyType == typeof(double) || columnProperty.PropertyType == typeof(double?))
                        {
                            columnProperty.PropertyInfo.SetValue(dataRow, cellValue.GetValue<double>());
                        }
                        else if (columnProperty.PropertyType == typeof(DateOnly) || columnProperty.PropertyType == typeof(DateOnly?))
                        {
                            columnProperty.PropertyInfo.SetValue(dataRow, DateOnly.FromDateTime(cellValue.GetDateTime()));
                        }
                        else if (columnProperty.PropertyType == typeof(DateTime) || columnProperty.PropertyType == typeof(DateTime?))
                        {
                            columnProperty.PropertyInfo.SetValue(dataRow, cellValue.GetValue<DateTime>());
                        }
                        else if (columnProperty.PropertyType == typeof(TimeOnly) || columnProperty.PropertyType == typeof(TimeOnly?))
                        {
                            columnProperty.PropertyInfo.SetValue(dataRow, TimeOnly.FromTimeSpan(cellValue.GetValue<TimeSpan>()));
                        }
                        else if (columnProperty.PropertyType == typeof(string))
                        {
                            columnProperty.PropertyInfo.SetValue(dataRow, cellValue.GetString());
                        }
                        else
                        {
                            throw new InvalidOperationException($"The property '{typeof(T).Name}.{columnProperty.PropertyInfo.Name}' is declared as '{columnProperty.PropertyType.Name}' which is not supported.");
                        }
                    }
                    catch (Exception e)
                    {
                        throw new InvalidCastException($"The conversion of the value {worksheet}!{cellValue} with type '{cellValue.DataType}' to property type '{columnProperty.PropertyType.Name}' failed.", e);
                    }
                }

                if (validationProblems.Count > 0)
                {
                    break;
                }
                
                data.Add(dataRow);
                rowIndex++;
            }
        }

        return new ConversionResult<T>(validationProblems, data);
    }
}