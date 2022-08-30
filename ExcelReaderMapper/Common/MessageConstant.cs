using ExcelReaderMapper.Infrastructure;
using ExcelReaderMapper.Model;

namespace ExcelReaderMapper.Common
{
    public static class MessageConstant
    {
        public static ILoggingModel InvalidTemplate => new LoggingModel()
        {
            Code = ErrorCodeEnum.InvalidTemplate,
            Message = "Invalid template in excel worksheet"
        };

        public static ILoggingModel MissingDataFirstRow =>
            new LoggingModel()
            {
                Code = ErrorCodeEnum.MissingDataFirstRow,
                Message =
                    "Missing data of the first row. The validation cannot be processed."
            };

        public static ILoggingModel MustInDateFormat(string? fieldName)
        {
            return new LoggingModel()
            {
                Code = ErrorCodeEnum.MustInDateFormat,
                Message = $"{fieldName} must be in date format.",
                ColumnName = fieldName
            };
        }

        public static ILoggingModel InvalidField(string? fieldName)
        {
            return new LoggingModel()
            {
                Code = ErrorCodeEnum.InvalidField,
                Message = $"{fieldName} is invalid",
                ColumnName = fieldName
            };
        }

        public static ILoggingModel MissingColumns(string? columnsName)
        {
            return new LoggingModel()
            {
                Code = ErrorCodeEnum.MissingColumns,
                Message =
                    $"The uploaded file is missing the column {columnsName}",
                ColumnName = columnsName
            };
        }

        public static ILoggingModel DuplicateColumn(string? columnsName)
        {
            return new LoggingModel()
            {
                Code = ErrorCodeEnum.DuplicateColumn,
                Message =
                    $"The following columns is duplicated in uploaded file: {columnsName}",
                ColumnName = columnsName
            };
        }

        public static ILoggingModel FieldMustBeNumeric(string? columnsName)
        {
            return new LoggingModel()
            {
                Code = ErrorCodeEnum.FieldMustBeNumeric,
                Message = $"{columnsName} must be numeric.",
                ColumnName = columnsName
            };
        }
    }
}