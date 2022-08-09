using System.Collections.Generic;
using System.IO;
using ExcelDataReader;
using ExcelReaderMapper.Common;
using ExcelReaderMapper.Infrastructure;
using ExcelReaderMapper.Model;

namespace ExcelReaderMapper.Service
{
    public class ExcelMapperService : IExcelMapperService
    {
        private IHeaderRowService? _headerRowService;

        private IValidateHeaderService? _validateHeaderService;

        public ExcelMapperService()
        {
        }

        public ExcelMapperService(IHeaderRowService? headerRowService)
        {
            HeaderRowService = headerRowService;
        }

        public ExcelMapperService(IValidateHeaderService? validateHeaderService)
        {
            ValidateHeaderService = validateHeaderService;
        }

        public IHeaderRowService? HeaderRowService
        {
            get
            {
                if (_headerRowService == null)
                {
                    _headerRowService = new HeaderRowService();
                }

                return _headerRowService;
            }
            set => _headerRowService = value;
        }

        public IValidateHeaderService? ValidateHeaderService
        {
            get
            {
                if (_validateHeaderService == null)
                {
                    _validateHeaderService = new ValidateHeaderService();
                }

                return _validateHeaderService;
            }
            set => _validateHeaderService = value;
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