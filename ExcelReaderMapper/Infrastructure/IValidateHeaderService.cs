using System.Collections.Generic;
using ExcelReaderMapper.Model;

namespace ExcelReaderMapper.Infrastructure
{
    public interface IValidateHeaderService
    {
        bool ValidateHeader<TExcelModel>(string[] headerNameExcel, List<ExcelColumnModel> excelColumnModels,
            out List<ILoggingModel> linesError);
    }
}