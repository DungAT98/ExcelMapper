using System.Collections.Generic;
using ExcelReaderMapper.Common;
using ExcelReaderMapper.Model;

namespace ExcelReaderMapper.Infrastructure
{
    public interface IExcelMapperService
    {
        IEnumerable<WorkSheetResult<TExcelModel>> GetDataFromExcel<TExcelModel>(byte[] content, int lineOffset = 1,
            int lengthOfHeader = 1, ParsingMethod parsingMethod = ParsingMethod.Reflection);

        WorkSheetResult<TExcelModel> GetDataFromExcel<TExcelModel>(byte[] content, int sheetIndex,
            int lineOffset = 1, int lengthOfHeader = 1, ParsingMethod parsingMethod = ParsingMethod.Reflection);
    }
}