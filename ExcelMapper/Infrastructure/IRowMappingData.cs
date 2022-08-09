using System.Collections.Generic;
using ExcelDataReader;
using ExcelReaderMapper.Model;

namespace ExcelReaderMapper.Infrastructure
{
    internal interface IRowMappingData
    {
        TModel MappingData<TModel>(IExcelDataReader reader, List<ExcelColumnModel> columnList,
            ref List<ILoggingModel> errorsList);
    }
}