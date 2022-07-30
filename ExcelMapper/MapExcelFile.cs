using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.IO;
using System.Linq;
using ExcelDataReader;
using ExcelMapper.Common;
using ExcelMapper.Infrastructure;
using ExcelMapper.Model;

namespace ExcelMapper
{
    public class MapExcelFile
    {
        public string[] DateTimeFormat { get; } = { };

        public List<IExcelResult<TExcelModel>> GetDataFromExcel<TExcelModel>(byte[] content, int lineOffset = 1)
        {
            var result = new List<IExcelResult<TExcelModel>>();
            using var memoryStream = new MemoryStream(content);
            using var reader = ExcelReaderFactory.CreateReader(memoryStream);
            var currentLine = lineOffset;
            var excelColumnsList = new List<ExcelColumnModel>();
            do
            {
                while (reader.Read())
                {
                    var rowData = ReadEntireRow(reader);
                    var isEmptyRow = IsRowEmpty(rowData);
                    if (isEmptyRow)
                    {
                        currentLine++;
                        continue;
                    }

                    if (currentLine == lineOffset)
                    {
                        var headerRow = ReadEntireRow(reader);
                        var isHeaderValid = ValidateHeader<TExcelModel>(headerRow, out var linesError);
                        if (!isHeaderValid)
                        {
                            reader.Close();
                            result.Add(new ExcelResult<TExcelModel>()
                            {
                                Errors = linesError,
                                ExcelModel = default!,
                                IsError = true,
                                LineNumber = lineOffset
                            });

                            return result;
                        }

                        excelColumnsList = MappingExcelColumnNumber<TExcelModel>(headerRow);
                    }
                    else
                    {
                        var data = GetDataFromCell<TExcelModel>(reader, excelColumnsList, out var errorsList);
                        result.Add(new ExcelResult<TExcelModel>()
                        {
                            ExcelModel = data,
                            LineNumber = currentLine,
                            Errors = errorsList,
                            IsError = errorsList.Count != 0
                        });
                    }

                    currentLine++;
                }

                // each sheet need to be reset the counter
                currentLine = lineOffset;
            } while (reader.NextResult());

            reader.Close();

            return result;
        }

        private bool IsRowEmpty(string[] data)
        {
            return data.All(string.IsNullOrWhiteSpace);
        }

        private string[] ReadEntireRow(IExcelDataReader reader)
        {
            var result = new string[reader.FieldCount];
            for (var i = 0; i < reader.FieldCount; i++)
                try
                {
                    result[i] = reader.GetValue(i)?.ToString()!;
                }
                catch
                {
                    result[i] = null!;
                }

            return result;
        }

        private bool ValidateHeader<TExcelModel>(string[] headerNameExcel, out List<ILoggingModel> linesError)
        {
            linesError = new List<ILoggingModel>();
            if (headerNameExcel.All(string.IsNullOrWhiteSpace))
            {
                linesError.Add(MessageConstant.MissingDataFirstRow);
                return false;
            }

            var attributeData = ReflectionHelper.GetAttributeValue<TExcelModel, DisplayAttribute>()
                .Where(n => !string.IsNullOrWhiteSpace(n.Name))
                .Select(n => n.Name)
                .ToList();
            var duplicateColumns = headerNameExcel.GroupBy(n => n)
                .Where(n => n.Count() > 1)
                .Select(n => n.Key)
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .ToList();
            var notValidColumn = attributeData.Except(headerNameExcel)
                .ToList();

            if (notValidColumn.Count > 0)
            {
                linesError.AddRange(notValidColumn.Select(MessageConstant.MissingColumns));
            }

            if (duplicateColumns.Count > 0)
            {
                linesError.AddRange(duplicateColumns.Select(MessageConstant.DuplicateColumn));
            }

            return duplicateColumns.Count == 0 && notValidColumn.Count == 0;
        }

        private List<ExcelColumnModel> MappingExcelColumnNumber<TModel>(string[] headerRow)
        {
            var result = new List<ExcelColumnModel>();
            var headerRowList = headerRow.ToList();
            var propertyInfos = ReflectionHelper.GetPropertyInfoWithAttribute<TModel, DisplayAttribute>();
            foreach (var propertyInfo in propertyInfos)
            {
                var attributeValue = ReflectionHelper.GetAttributeValue<DisplayAttribute>(propertyInfo)
                    .FirstOrDefault()?.Name;
                if (attributeValue == null)
                {
                    continue;
                }

                var columnMatched = headerRowList.ToList().IndexOf(attributeValue);
                if (columnMatched == -1)
                {
                    continue;
                }

                var entity = new ExcelColumnModel()
                {
                    PropertyInfo = propertyInfo,
                    ColumnNumber = columnMatched,
                    Name = attributeValue
                };

                result.Add(entity);
            }

            return result;
        }

        private TModel GetDataFromCell<TModel>(IExcelDataReader reader, List<ExcelColumnModel> columnList,
            out List<ILoggingModel> errorsList)
        {
            errorsList = new List<ILoggingModel>();
            var currentObject = Activator.CreateInstance<TModel>();
            foreach (var excelColumnModel in columnList)
                try
                {
                    var isNullable = Nullable.GetUnderlyingType(excelColumnModel.PropertyInfo.PropertyType) != null;
                    if (isNullable &&
                        string.IsNullOrWhiteSpace(reader.GetValue(excelColumnModel.ColumnNumber)?.ToString()))
                    {
                        excelColumnModel.PropertyInfo.SetValue(currentObject, null);
                    }
                    else
                    {
                        var t = Nullable.GetUnderlyingType(excelColumnModel.PropertyInfo.PropertyType) ??
                                excelColumnModel.PropertyInfo.PropertyType;
                        if (t == typeof(DateTime) || t == typeof(DateTime?))
                        {
                            DateTime dateTimeValue;
                            // type is DateTime
                            try
                            {
                                dateTimeValue = reader.GetDateTime(excelColumnModel.ColumnNumber);
                            }
                            catch
                            {
                                // not a datetime field, parse it using string
                                var canCastDatetime = DateTime.TryParseExact(
                                    reader.GetValue(excelColumnModel.ColumnNumber)?.ToString(), DateTimeFormat,
                                    CultureInfo.CurrentCulture, DateTimeStyles.None, out dateTimeValue);

                                if (!canCastDatetime)
                                {
                                    errorsList.Add(MessageConstant.MustInDateFormat(excelColumnModel.Name));
                                    continue;
                                }
                            }

                            excelColumnModel.PropertyInfo.SetValue(currentObject, dateTimeValue);
                        }
                        else if (t == typeof(decimal) || t == typeof(decimal?))
                        {
                            var canParseDecimal =
                                decimal.TryParse(reader.GetValue(excelColumnModel.ColumnNumber)?.ToString(),
                                    out var decimalValue);
                            if (!canParseDecimal)
                            {
                                errorsList.Add(MessageConstant.FieldMustBeNumeric(excelColumnModel.Name));
                                continue;
                            }

                            excelColumnModel.PropertyInfo.SetValue(currentObject, decimalValue);
                        }
                        else if (t == typeof(double) || t == typeof(double?))
                        {
                            var canParseDecimal =
                                double.TryParse(reader.GetValue(excelColumnModel.ColumnNumber)?.ToString(),
                                    out var decimalValue);
                            if (!canParseDecimal)
                            {
                                errorsList.Add(MessageConstant.FieldMustBeNumeric(excelColumnModel.Name));
                                continue;
                            }

                            excelColumnModel.PropertyInfo.SetValue(currentObject, decimalValue);
                        }
                        else
                        {
                            var value = reader.GetValue(excelColumnModel.ColumnNumber);
                            var safeValue = TypeDescriptor.GetConverter(t)
                                .ConvertFromInvariantString(value?.ToString());
                            excelColumnModel.PropertyInfo.SetValue(currentObject, safeValue);
                        }
                    }
                }
                catch
                {
                    errorsList.Add(MessageConstant.InvalidField(excelColumnModel.Name));
                }

            return currentObject;
        }
    }
}