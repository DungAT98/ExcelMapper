using System.Text;
using ExcelReaderMapper.Service;

namespace ExcelReaderMapper.Testing.Tests;

public class BaseUnitTest
{
    protected readonly ExcelMapperService _excelMapperService;

    public BaseUnitTest()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        _excelMapperService = new ExcelMapperService();
    }
}