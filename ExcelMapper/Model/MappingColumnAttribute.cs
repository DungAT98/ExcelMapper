using System;

namespace ExcelReaderMapper.Model
{
    [AttributeUsage(AttributeTargets.Property)]
    public class MappingColumnAttribute : Attribute
    {
        public string Name { get; set; } = string.Empty;

        public string[]? CustomFormat { get; set; }
    }
}