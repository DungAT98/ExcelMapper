using System.Collections.Generic;

namespace ExcelReaderMapper.Infrastructure
{
    public interface IExcelRowResult<TExcelModel>
    {
        public TExcelModel ExcelModel { get; set; }

        public int LineNumber { get; set; }

        public bool IsError { get; set; }

        public List<ILoggingModel>? Errors { get; set; }
    }
}