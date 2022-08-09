using System;
using System.Globalization;
using System.Linq;
using ExcelDataReader;

namespace ExcelReaderMapper.Common
{
    internal static class ExcelHelper
    {
        public static bool IsRowEmpty(string[] data)
        {
            return data.All(string.IsNullOrWhiteSpace);
        }

        public static string[] ReadEntireRow(IExcelDataReader reader)
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

        public static bool TryConvertToDateTime(object value, string[]? format, out object result)
        {
            result = null!;

            if (value is string stringValue)
            {
                // string to DateTime
                if (format != null && format.Length != 0)
                {
                    if (DateTime.TryParseExact(stringValue, format, CultureInfo.CurrentCulture,
                            DateTimeStyles.AllowWhiteSpaces, out var dateTime2))
                    {
                        result = dateTime2;
                        return true;
                    }
                }

                if (DateTime.TryParse(stringValue, CultureInfo.CurrentCulture, DateTimeStyles.AllowWhiteSpaces,
                        out var dateTime))
                {
                    result = dateTime;
                    return true;
                }

                return false;
            }

            if (value is DateTime dateTimeValue)
            {
                result = dateTimeValue;

                return true;
            }

            return false;
        }
    }
}