using System.Collections.Generic;

namespace ExcelReaderMapper.Infrastructure
{
    public interface IValidateHeaderService
    {
        bool ValidateHeader<TExcelModel>(string[] headerNameExcel, out List<ILoggingModel> linesError);
    }
}