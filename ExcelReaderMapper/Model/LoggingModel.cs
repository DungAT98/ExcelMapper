using ExcelReaderMapper.Common;
using ExcelReaderMapper.Infrastructure;

namespace ExcelReaderMapper.Model
{
    public class LoggingModel : ILoggingModel
    {
        public ErrorCodeEnum Code { get; set; }

        public string? Message { get; set; }
    }
}