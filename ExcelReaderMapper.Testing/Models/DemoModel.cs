using ExcelReaderMapper.Model;

namespace ExcelReaderMapper.Testing.Models;

public class DemoModel
{
    [MappingColumn(Name = "No")]
    public int No { get; set; }

    [MappingColumn(Name = "First Name")]
    public string? FirstName { get; set; }

    [MappingColumn(Name = "Last Name")]
    public string? LastName { get; set; }

    [MappingColumn(Name = "Address")]
    public string? Address { get; set; }

    [MappingColumn(Name = "Id")]
    public Guid Id { get; set; }

    [MappingColumn(Name = "BirthDate")]
    public DateTime BirthDate { get; set; }
}