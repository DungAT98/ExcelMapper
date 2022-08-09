using System.Collections.Generic;
using System.IO;
using ExcelDataReader;
using ExcelReaderMapper.Common;
using ExcelReaderMapper.Infrastructure;
using ExcelReaderMapper.Model;

namespace ExcelReaderMapper.Service
{
    public class MultiLevelExcelMapperService : ExcelMapperServiceBase, IExcelMapperService
    {
        public MultiLevelExcelMapperService()
        {
        }

        public MultiLevelExcelMapperService(IHeaderRowService? headerRowService) : base(headerRowService)
        {
        }

        public MultiLevelExcelMapperService(IValidateHeaderService? validateHeaderService) : base(validateHeaderService)
        {
        }

        public MultiLevelExcelMapperService(IHeaderRowService? headerRowService,
            IValidateHeaderService? validateHeaderService) : base(headerRowService, validateHeaderService)
        {
        }

        public List<IExcelResult<TExcelModel>> GetDataFromExcel<TExcelModel>(byte[] content, int lineOffset = 1,
            ParsingMethod parsingMethod = ParsingMethod.Reflection)
        {
            var result = new List<IExcelResult<TExcelModel>>();
            using var memoryStream = new MemoryStream(content);
            using var reader = ExcelReaderFactory.CreateReader(memoryStream);
            var currentLine = 1;
            var excelColumnsList = new List<ExcelColumnModel>();
            do
            {
                while (reader.Read())
                {
                    // skip all above lines until got the lineOffset
                    if (currentLine < lineOffset)
                    {
                        continue;
                    }

                    var rowData = ExcelHelper.ReadEntireRow(reader);
                    var isEmptyRow = ExcelHelper.IsRowEmpty(rowData);
                    if (isEmptyRow)
                    {
                        currentLine++;
                        continue;
                    }

                    if (currentLine == lineOffset)
                    {
                        var headerRow = ExcelHelper.ReadEntireRow(reader);
                        var isHeaderValid =
                            ValidateHeaderService!.ValidateHeader<TExcelModel>(headerRow, out var linesError);
                        if (!isHeaderValid)
                        {
                            reader.Close();
                            result.Add(new ExcelResult<TExcelModel>
                            {
                                Errors = linesError,
                                ExcelModel = default!,
                                IsError = true,
                                LineNumber = lineOffset
                            });

                            return result;
                        }

                        excelColumnsList = HeaderRowService!.MappingExcelColumnNumber<TExcelModel>(headerRow);
                    }
                    else
                    {
                        var data = GetDataFromCell<TExcelModel>(reader, excelColumnsList, out var errorsList,
                            parsingMethod);
                        result.Add(new ExcelResult<TExcelModel>
                        {
                            ExcelModel = data,
                            LineNumber = currentLine,
                            Errors = errorsList,
                            IsError = errorsList.Count != 0
                        });
                    }

                    currentLine++;
                }

                // each sheet need to be reset the counter
                currentLine = lineOffset;
            } while (reader.NextResult());

            reader.Close();

            return result;
        }
    }
}