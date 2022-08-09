using System.Collections.Generic;
using System.Linq;
using ExcelReaderMapper.Common;
using ExcelReaderMapper.Infrastructure;
using ExcelReaderMapper.Model;

namespace ExcelReaderMapper.Service
{
    public class ValidateHeaderService : IValidateHeaderService
    {
        public bool ValidateHeader<TExcelModel>(string[] headerNameExcel, out List<ILoggingModel> linesError)
        {
            linesError = new List<ILoggingModel>();
            if (headerNameExcel.All(string.IsNullOrWhiteSpace))
            {
                linesError.Add(MessageConstant.MissingDataFirstRow);
                return false;
            }

            var attributeData = ReflectionHelper.GetAttributeValue<TExcelModel, MappingColumnAttribute>()
                .Where(n => !string.IsNullOrWhiteSpace(n.Name))
                .Select(n => n.Name)
                .ToList();
            var duplicateColumns = headerNameExcel.GroupBy(n => n)
                .Where(n => n.Count() > 1)
                .Select(n => n.Key)
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .ToList();
            var notValidColumn = attributeData.Except(headerNameExcel)
                .ToList();

            if (notValidColumn.Count > 0) linesError.AddRange(notValidColumn.Select(MessageConstant.MissingColumns));

            if (duplicateColumns.Count > 0)
                linesError.AddRange(duplicateColumns.Select(MessageConstant.DuplicateColumn));

            return duplicateColumns.Count == 0 && notValidColumn.Count == 0;
        }
    }
}