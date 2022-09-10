using System.Collections.Generic;
using System.Linq;

namespace ExcelReaderMapper.Model
{
    public class WorkSheetResult<TExcelModel>
    {
        public string SheetName { get; set; }

        public int SheetNumber { get; set; }

        public List<ExcelRowResult<TExcelModel>> RowResults { get; set; } =
            new List<ExcelRowResult<TExcelModel>>();

        public bool IsError => RowResults.Count(n => n.IsError) > 0;
    }
}