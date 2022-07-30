using System.Reflection;

namespace ExcelMapper.Model
{
    public class ExcelColumnModel
    {
        public string? Name { get; set; }

        public int ColumnNumber { get; set; }

        public PropertyInfo? PropertyInfo { get; set; }
    }
}