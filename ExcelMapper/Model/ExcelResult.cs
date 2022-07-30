using System.Collections.Generic;
using ExcelMapper.Infrastructure;

namespace ExcelMapper.Model
{
    public class ExcelResult<TExcelModel> : IExcelResult<TExcelModel>
    {
        public List<ILoggingModel>? Errors { get; set; }

        public TExcelModel ExcelModel { get; set; } = default!;

        public int LineNumber { get; set; }

        public bool IsError { get; set; }
    }
}