using System.Linq;
using ExcelDataReader;

namespace ExcelMapper.Common
{
    public static class ExcelHelper
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
    }
}