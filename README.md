# ExcelReaderMapper

## Lightweight excel reader for .NET application

[![Build project](https://github.com/DungAT98/ExcelMapper/actions/workflows/builld.yml/badge.svg)](https://github.com/DungAT98/ExcelMapper/actions/workflows/builld.yml) [![publish to nuget](https://github.com/DungAT98/ExcelMapper/actions/workflows/dotnet.yml/badge.svg)](https://github.com/DungAT98/ExcelMapper/actions/workflows/dotnet.yml)

This use [ExcelDataReader](https://www.nuget.org/packages/ExcelDataReader/) to read the excel file. It will read line
one by one. So it optimize the memory

## Installation

Download package from NuGet [ExcelReaderMapper](https://www.nuget.org/packages/ExcelReaderMapper)

## Note for .NET CORE

By default, ExcelDataReader throws a NotSupportedException "No data is available for encoding 1252." on .NET Core.

To fix, add a dependency to the package System.Text.Encoding.CodePages and then add code to register the code page
provider during application initialization (f.ex in Startup.cs):

```
System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
```

This is required to parse strings in binary BIFF2-5 Excel documents encoded with DOS-era code pages. These encodings are
registered by default in the full .NET Framework, but not on .NET Core.

# Using

Create a model to mapping with datatype with DisplayAttribute to declare the excel column name

```
public class ExcelModel
{
    [MappingColumn(Name = "FIRST_NAME")]
    public string FirstName { get; set; }
    
    [MappingColumn(Name = "LAST_NAME")]
    public string LastNane { get; set; }
    
    [MappingColumn(Name = "AGE", CustomFormat = { "dd-mm-yyyy" })]
    public DateTime DateOfBirth { get; set; }
}
```

Call the excelmapper object

```
var fileContent = File.ReadAllBytes("Example.xlsx");
var excelMapper = new ExcelMapperService();
var result = excelMapper.GetDataFromExcel<ExcelModel>(fileContent);
```

## Properties

lineOfOffset: It will offset the lines that started reading the excel file (Default = 1)
parsingMethod: Additional parsing method (not to use reflection too much) - refer ParsingMethod enum. (Experimental)