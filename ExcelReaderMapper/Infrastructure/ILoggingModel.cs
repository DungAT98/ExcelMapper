using ExcelReaderMapper.Common;

namespace ExcelReaderMapper.Infrastructure
{
    public interface ILoggingModel
    {
        public ErrorCodeEnum Code { get; set; }

        public string? Message { get; set; }

        public string? ColumnName { get; set; }
    }
}