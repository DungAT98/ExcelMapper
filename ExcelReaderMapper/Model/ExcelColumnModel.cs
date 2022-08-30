using System.Reflection;

namespace ExcelReaderMapper.Model
{
    public class ExcelColumnModel
    {
        public string Name { get; set; } = string.Empty;

        public int ColumnNumber { get; set; }

        public MappingColumnAttribute MappingColumnAttribute { get; set; }

        public PropertyInfo PropertyInfo { get; set; }
    }
}