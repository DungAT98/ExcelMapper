using ExcelMapper.Common;
using ExcelMapper.Infrastructure;

namespace ExcelMapper.Model
{
    public class LoggingModel : ILoggingModel
    {
        public ErrorCodeEnum Code { get; set; }

        public string? Message { get; set; }
    }
}