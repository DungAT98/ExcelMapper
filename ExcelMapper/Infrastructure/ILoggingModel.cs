using ExcelMapper.Common;

namespace ExcelMapper.Infrastructure
{
    public interface ILoggingModel
    {
        public ErrorCodeEnum Code { get; set; }

        public string? Message { get; set; }
    }
}