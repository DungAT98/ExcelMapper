 using System.Collections.Generic;

namespace ExcelMapper.Infrastructure
{
    public interface IMapExcelFile
    {
        List<IExcelResult<TExcelModel>> GetDataFromExcel<TExcelModel>(byte[] content, int lineOffset = 1);
    }
}