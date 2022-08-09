using ExcelReaderMapper.Model;

namespace ExcelReaderMapper.Testing.Models;

public class DemoModel
{
    [MappingColumn(Name = "No")]
    public int No { get; set; }

    [MappingColumn(Name = "Name")]
    public string? Name { get; set; }

    [MappingColumn(Name = "Address")]
    public string? Address { get; set; }
}