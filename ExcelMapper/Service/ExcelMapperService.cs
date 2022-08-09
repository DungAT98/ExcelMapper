using System.Collections.Generic;
using System.IO;
using ExcelDataReader;
using ExcelReaderMapper.Common;
using ExcelReaderMapper.Factory;
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

        public List<IExcelResult<TExcelModel>> GetDataFromExcel<TExcelModel>(byte[] content, int lineOffset = 1,
            ParsingMethod parsingMethod = ParsingMethod.Reflection)
        {
            var result = new List<IExcelResult<TExcelModel>>();
            using var memoryStream = new MemoryStream(content);
            using var reader = ExcelReaderFactory.CreateReader(memoryStream);
            var currentLine = lineOffset;
            var excelColumnsList = new List<ExcelColumnModel>();
            do
            {
                while (reader.Read())
                {
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

        private TModel GetDataFromCell<TModel>(IExcelDataReader reader, List<ExcelColumnModel> columnList,
            out List<ILoggingModel> errorsList, ParsingMethod parsingMethod)
        {
            errorsList = new List<ILoggingModel>();
            var parsingFactory = GetDataFromCellFactory.GetDataFromCell(parsingMethod);
            var result = parsingFactory.MappingData<TModel>(reader, columnList, ref errorsList);

            return result;
        }
    }
}