using System.Collections.Generic;
using ExcelReaderMapper.Common;

namespace ExcelReaderMapper.Infrastructure
{
    public interface IExcelMapperService
    {
        List<IExcelResult<TExcelModel>> GetDataFromExcel<TExcelModel>(byte[] content, int lineOffset = 1,
            ParsingMethod parsingMethod = ParsingMethod.Reflection);
    }
}