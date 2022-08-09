using System.Collections.Generic;
using ExcelReaderMapper.Model;

namespace ExcelReaderMapper.Infrastructure
{
    public interface IHeaderRowService
    {
        List<ExcelColumnModel> MappingExcelColumnNumber<TModel>(string[] headerRow);
    }
}