using System;
using System.Collections.Generic;
using System.Globalization;
using ExcelDataReader;
using ExcelReaderMapper.Common;
using ExcelReaderMapper.Extensions;
using ExcelReaderMapper.Infrastructure;
using ExcelReaderMapper.Model;

namespace ExcelReaderMapper.Factory
{
    internal class RowMappingParsing : IRowMappingData
    {
        public TModel MappingData<TModel>(IExcelDataReader reader, List<ExcelColumnModel> columnList,
            ref List<ILoggingModel> errorsList)
        {
            var model = Activator.CreateInstance<TModel>();
            foreach (var excelColumnModel in columnList)
            {
                var isConverted = TryConvertType(reader, excelColumnModel, out var result);
                if (isConverted)
                {
                    excelColumnModel.PropertyInfo.SetValue(model, result);
                }
                else
                {
                    errorsList.Add(MessageConstant.InvalidField(excelColumnModel.MappingColumnAttribute?.Name));
                }
            }

            return model;
        }

        private bool TryConvertType(IExcelDataReader reader, ExcelColumnModel columnModel, out object? result)
        {
            result = null;
            var isNullable = false;
            var targetType = columnModel.PropertyInfo.PropertyType;
            // get base type of nullable type
            if (Nullable.GetUnderlyingType(columnModel.PropertyInfo.PropertyType) != null)
            {
                isNullable = true;
                targetType = Nullable.GetUnderlyingType(columnModel.PropertyInfo.PropertyType);
            }

            if (targetType == null)
            {
                return false;
            }

            object value;
            try
            {
                value = reader.GetDateTime(columnModel.ColumnNumber);
            }
            catch
            {
                value = reader.GetValue(columnModel.ColumnNumber)?.ToString() ?? string.Empty;
            }

            if (targetType == typeof(DateTime))
            {
                var isConverted = ExcelHelper.TryConvertToDateTime(value,
                    columnModel.MappingColumnAttribute?.CustomFormat, out result);
                return isConverted;
            }

            if (value is string stringValue)
            {
                if (targetType == typeof(string))
                {
                    result = stringValue;
                    return true;
                }

                if (targetType.IsNumeric() &&
                    double.TryParse(stringValue, NumberStyles.Any, null, out var doubleResult))
                {
                    result = Convert.ChangeType(doubleResult, targetType);
                    return true;
                }

                if (targetType.IsEnum)
                {
                    result = Enum.Parse(targetType, stringValue, true);
                    return true;
                }

                if (targetType == typeof(Guid))
                {
                    var parsed = Guid.TryParse(stringValue, out var guidResult);
                    result = guidResult;
                    return parsed;
                }

                // Ensure we are not throwing exception and just read a null for nullable property.
                if (isNullable && string.IsNullOrWhiteSpace(stringValue))
                {
                    return true;
                }
            }

            try
            {
                result = Convert.ChangeType(value, targetType);
            }
            catch
            {
                return false;
            }

            return true;
        }
    }
}