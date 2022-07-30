using ExcelMapper.Infrastructure;
using ExcelMapper.Model;

namespace ExcelMapper.Common
{
    public static class MessageConstant
    {
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
                Message = $"{fieldName} must be in date format."
            };
        }

        public static ILoggingModel InvalidField(string? fieldName)
        {
            return new LoggingModel()
            {
                Code = ErrorCodeEnum.InvalidField,
                Message = $"{fieldName} is invalid"
            };
        }

        public static ILoggingModel MissingColumns(string? columnsName)
        {
            return new LoggingModel()
            {
                Code = ErrorCodeEnum.MissingColumns,
                Message =
                    $"The uploaded file is missing the column {columnsName}"
            };
        }

        public static ILoggingModel DuplicateColumn(string? columnsName)
        {
            return new LoggingModel()
            {
                Code = ErrorCodeEnum.DuplicateColumn,
                Message =
                    $"The following columns is duplicated in uploaded file: {columnsName}"
            };
        }

        public static ILoggingModel FieldMustBeNumeric(string? columnsName)
        {
            return new LoggingModel()
            {
                Code = ErrorCodeEnum.FieldMustBeNumeric,
                Message = $"{columnsName} must be numeric."
            };
        }
    }
}