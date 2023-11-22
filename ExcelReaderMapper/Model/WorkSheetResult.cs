using System.Collections.Generic;
using System.Linq;

namespace ExcelReaderMapper.Model
{
    public class WorkSheetResult<TExcelModel> : WorksheetBaseInformation
    {
        public List<ExcelRowResult<TExcelModel>> RowResults { get; set; } =
            new List<ExcelRowResult<TExcelModel>>();

        public bool IsError => RowResults.Count(n => n.IsError) > 0;

        public Dictionary<string, int> HeaderRow { get; set; } = new Dictionary<string, int>();
    }
}