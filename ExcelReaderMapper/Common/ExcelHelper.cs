using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelDataReader;
using ExcelReaderMapper.Extensions;

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
            var result = new List<string>();
            for (var i = 0; i < reader.FieldCount; i++)
            {
                try
                {
                    result.Add(reader.GetValue(i)?.ToString());
                }
                catch
                {
                    result.Add(null);
                }
            }

            while (string.IsNullOrWhiteSpace(result.TakeLast(1).FirstOrDefault()) && result.Count > 0)
            {
                result.RemoveAt(result.Count - 1);
            }

            return result.ToArray();
        }

        public static bool TryConvertToDateTime(object value, string[] format, out object result)
        {
            result = null;

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