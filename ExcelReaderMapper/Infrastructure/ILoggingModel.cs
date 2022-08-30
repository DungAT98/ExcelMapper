using ExcelReaderMapper.Common;

namespace ExcelReaderMapper.Infrastructure
{
    public interface ILoggingModel
    {
        ErrorCodeEnum Code { get; set; }

        string Message { get; set; }

        string ColumnName { get; set; }
    }
}