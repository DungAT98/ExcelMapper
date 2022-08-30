using System.Collections.Generic;
using System.IO;
using ExcelDataReader;
using ExcelReaderMapper.Common;
using ExcelReaderMapper.Infrastructure;
using ExcelReaderMapper.Model;

namespace ExcelReaderMapper.Service
{
    public class ExcelMapperService : ExcelMapperServiceBase, IExcelMapperService
    {
        public ExcelMapperService()
        {
        }

        public ExcelMapperService(IHeaderRowService? headerRowService) : base(headerRowService)
        {
        }

        public ExcelMapperService(IValidateHeaderService? validateHeaderService) : base(validateHeaderService)
        {
        }

        public ExcelMapperService(IHeaderRowService? headerRowService, IValidateHeaderService? validateHeaderService) :
            base(headerRowService, validateHeaderService)
        {
        }

        public IEnumerable<WorkSheetResult<TExcelModel>> GetDataFromExcel<TExcelModel>(byte[] content,
            int lineOffset = 1, int lengthOfHeader = 1, ParsingMethod parsingMethod = ParsingMethod.Reflection)
        {
            var sheetIndex = 1;
            var result = new List<WorkSheetResult<TExcelModel>>();
            using var memoryStream = new MemoryStream(content);
            using var reader = ExcelReaderFactory.CreateReader(memoryStream);
            do
            {
                var workSheetResult =
                    MappingExcelInWorksheet<TExcelModel>(reader, lineOffset, sheetIndex, lengthOfHeader,
                        parsingMethod);
                // each sheet need to be reset the counter
                result.Add(workSheetResult);
                sheetIndex++;
            } while (reader.NextResult());

            reader.Close();

            return result;
        }

        public WorkSheetResult<TExcelModel>? GetDataFromExcel<TExcelModel>(byte[] content, int sheetIndex,
            int lineOffset = 1, int lengthOfHeader = 1, ParsingMethod parsingMethod = ParsingMethod.Reflection)
        {
            var currentSheetIndex = 1;
            using var memoryStream = new MemoryStream(content);
            using var reader = ExcelReaderFactory.CreateReader(memoryStream);
            do
            {
                if (sheetIndex != currentSheetIndex)
                {
                    currentSheetIndex++;
                    continue;
                }

                var result = MappingExcelInWorksheet<TExcelModel>(reader, lineOffset, sheetIndex, lengthOfHeader,
                    parsingMethod);
                reader.Close();

                return result;
                // each sheet need to be reset the counter
            } while (reader.NextResult());

            reader.Close();

            return null;
        }
    }
}