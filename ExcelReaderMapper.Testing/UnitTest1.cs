using System.Text;
using ExcelReaderMapper.Common;
using ExcelReaderMapper.Service;
using ExcelReaderMapper.Testing.Models;

namespace ExcelReaderMapper.Testing;

public class UnitTest1
{
    [Fact]
    public void Test1()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var fileContent = File.ReadAllBytes("Book1.xlsx");
        var excelMapping = new ExcelMapperService();
        var result = excelMapping.GetDataFromExcel<DemoModel>(fileContent, 1, ParsingMethod.NormalCase);
    }
}