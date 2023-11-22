using ExcelReaderMapper.Common;
using ExcelReaderMapper.Infrastructure;
using ExcelReaderMapper.Model;
using ExcelReaderMapper.Service;
using Xunit;

namespace TestExcelMultipleHeaderRow;

public class ExcelTest
{
    [Fact]
    public void TestMappingExcel()
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        var fileContent = File.ReadAllBytes("D:/FILE_IMPORT.xlsx");
        var mapper = new ExcelMapperService(new AlternativeValidateHeaderService());
        var header = mapper.GetHeaderRows(fileContent, 2, 2);
        var data = mapper.GetDataFromExcel<ExcelModel>(fileContent, lineOffset: 2, lengthOfHeader: 2);
    }
}

public class AlternativeValidateHeaderService : IValidateHeaderService
{
    public bool ValidateHeader<TExcelModel>(string[] headerNameExcel, List<ExcelColumnModel> excelColumnModels,
        out List<ILoggingModel> linesError)
    {
        linesError = new List<ILoggingModel>();
        if (excelColumnModels.Count == 0)
        {
            linesError.Add(MessageConstant.MissingDataFirstRow);
            return false;
        }

        if (headerNameExcel.All(string.IsNullOrWhiteSpace))
        {
            linesError.Add(MessageConstant.MissingDataFirstRow);
            return false;
        }

        var attributeData = ReflectionHelper.GetAttributeValue<TExcelModel, MappingColumnAttribute>()
            .Where(n => !string.IsNullOrWhiteSpace(n.Name))
            .Select(n => n.Name)
            .ToList();
        var notValidColumn = attributeData.Except(headerNameExcel)
            .ToList();

        if (notValidColumn.Count > 0) linesError.AddRange(notValidColumn.Select(MessageConstant.MissingColumns));

        return notValidColumn.Count == 0;
    }
}

public class ExcelModel
{
    [MappingColumn(Index = 1)]
    public string Line { get; set; }

    [MappingColumn(Index = 2)]
    public string Kishu { get; set; }

    [MappingColumn(Index = 3)]
    public int? NumberOfSet { get; set; }

    [MappingColumn(Index = 4)]
    public string To1 { get; set; }

    [MappingColumn(Index = 5)]
    public string To2 { get; set; }

    [MappingColumn(Index = 6)]
    public string To3 { get; set; }

    [MappingColumn(Index = 20)]
    public string Qty_To1 { get; set; }
}