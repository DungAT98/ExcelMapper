using System.Collections.Generic;
using ExcelReaderMapper.Infrastructure;

namespace ExcelReaderMapper.Model
{
    public class ExcelRowResult<TExcelModel> : IExcelRowResult<TExcelModel>
    {
        public List<ILoggingModel>? Errors { get; set; }

        public TExcelModel ExcelModel { get; set; } = default!;

        public int LineNumber { get; set; }

        public bool IsError { get; set; }
    }
}