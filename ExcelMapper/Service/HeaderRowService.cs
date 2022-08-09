using System.Collections.Generic;
using System.Linq;
using ExcelReaderMapper.Common;
using ExcelReaderMapper.Infrastructure;
using ExcelReaderMapper.Model;

namespace ExcelReaderMapper.Service
{
    public class HeaderRowService : IHeaderRowService
    {
        public List<ExcelColumnModel> MappingExcelColumnNumber<TModel>(string[] headerRow)
        {
            var result = new List<ExcelColumnModel>();
            var headerRowList = headerRow.ToList();
            var propertyInfos = ReflectionHelper.GetPropertyInfoWithAttribute<TModel, MappingColumnAttribute>();
            foreach (var propertyInfo in propertyInfos)
            {
                var attributeValue = ReflectionHelper.GetAttributeValue<MappingColumnAttribute>(propertyInfo)
                    .FirstOrDefault();
                if (attributeValue == null)
                {
                    continue;
                }

                var columnMatched = headerRowList.ToList().IndexOf(attributeValue.Name);
                if (columnMatched == -1)
                {
                    continue;
                }

                var entity = new ExcelColumnModel
                {
                    PropertyInfo = propertyInfo,
                    ColumnNumber = columnMatched,
                    MappingColumnAttribute = attributeValue,
                    Name = attributeValue.Name
                };

                result.Add(entity);
            }

            return result;
        }
    }
}