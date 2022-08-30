using System.Collections.Generic;

namespace ExcelReaderMapper.Infrastructure
{
    public interface IExcelRowResult<TExcelModel>
    {
        TExcelModel ExcelModel { get; set; }

        int LineNumber { get; set; }

        bool IsError { get; set; }

        List<ILoggingModel> Errors { get; set; }
    }
}