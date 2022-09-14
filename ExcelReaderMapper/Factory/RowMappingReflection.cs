using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using ExcelDataReader;
using ExcelReaderMapper.Common;
using ExcelReaderMapper.Infrastructure;
using ExcelReaderMapper.Model;

namespace ExcelReaderMapper.Factory
{
    internal class RowMappingReflection : IRowMappingData
    {
        public TModel MappingData<TModel>(IExcelDataReader reader, List<ExcelColumnModel> columnList,
            ref List<ILoggingModel> errorsList)
        {
            var currentObject = Activator.CreateInstance<TModel>();
            foreach (var excelColumnModel in columnList)
            {
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
                                if (excelColumnModel.MappingColumnAttribute?.CustomFormat != null &&
                                    excelColumnModel.MappingColumnAttribute?.CustomFormat.Length > 0)
                                {
                                    errorsList.Add(
                                        MessageConstant.MustInDateFormat(excelColumnModel.MappingColumnAttribute
                                            ?.Name));
                                    continue;
                                }

                                // not a datetime field, parse it using string
                                var canCastDatetime = DateTime.TryParseExact(
                                    reader.GetValue(excelColumnModel.ColumnNumber)?.ToString(),
                                    excelColumnModel.MappingColumnAttribute?.CustomFormat,
                                    CultureInfo.CurrentCulture, DateTimeStyles.None, out dateTimeValue);

                                if (!canCastDatetime)
                                {
                                    errorsList.Add(
                                        MessageConstant.MustInDateFormat(excelColumnModel.MappingColumnAttribute
                                            ?.Name));
                                    continue;
                                }
                            }

                            excelColumnModel.PropertyInfo.SetValue(currentObject, dateTimeValue);
                        }
                        else if (t == typeof(decimal) || t == typeof(decimal?))
                        {
                            var canParseDecimal =
                                decimal.TryParse(reader.GetValue(excelColumnModel.ColumnNumber)?.ToString(),
                                    NumberStyles.AllowExponent | NumberStyles.AllowDecimalPoint, null,
                                    out var decimalValue);
                            if (!canParseDecimal)
                            {
                                errorsList.Add(
                                    MessageConstant.FieldMustBeNumeric(excelColumnModel.MappingColumnAttribute?.Name));
                                continue;
                            }

                            excelColumnModel.PropertyInfo.SetValue(currentObject, decimalValue);
                        }
                        else if (t == typeof(double) || t == typeof(double?))
                        {
                            var canParseDecimal =
                                double.TryParse(reader.GetValue(excelColumnModel.ColumnNumber)?.ToString(),
                                    NumberStyles.AllowExponent | NumberStyles.AllowDecimalPoint, null,
                                    out var decimalValue);
                            if (!canParseDecimal)
                            {
                                errorsList.Add(
                                    MessageConstant.FieldMustBeNumeric(excelColumnModel.MappingColumnAttribute?.Name));
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
                    errorsList.Add(MessageConstant.InvalidField(excelColumnModel.MappingColumnAttribute?.Name));
                }
            }

            return currentObject;
        }
    }
}