using ExcelReaderMapper.Model;

namespace ExcelReaderMapper.Testing.Models;

public class ModelWithIndex
{
    [MappingColumn(Name = "No", Index = 0)]
    public int No { get; set; }

    [MappingColumn(Name = "First Name", Index = 1)]
    public string? FirstName { get; set; }

    [MappingColumn(Name = "Last Name", Index = 2)]
    public string? LastName { get; set; }

    [MappingColumn(Name = "Address", Index = 3)]
    public string? Address { get; set; }

    [MappingColumn(Name = "Id", Index = 4)]
    public Guid? Id { get; set; }

    [MappingColumn(Name = "BirthDate", Index = 5)]
    public DateTime BirthDate { get; set; }
}