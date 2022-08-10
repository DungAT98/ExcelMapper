using System.Collections.Generic;
using System.Linq;
using ExcelDataReader;
using ExcelReaderMapper.Common;
using ExcelReaderMapper.Factory;
using ExcelReaderMapper.Infrastructure;
using ExcelReaderMapper.Model;

namespace ExcelReaderMapper.Service
{
    public abstract class ExcelMapperServiceBase
    {
        private IHeaderRowService? _headerRowService;

        private IValidateHeaderService? _validateHeaderService;

        public ExcelMapperServiceBase()
        {
        }

        public ExcelMapperServiceBase(IHeaderRowService? headerRowService)
        {
            HeaderRowService = headerRowService;
        }

        public ExcelMapperServiceBase(IValidateHeaderService? validateHeaderService)
        {
            ValidateHeaderService = validateHeaderService;
        }

        public ExcelMapperServiceBase(IHeaderRowService? headerRowService,
            IValidateHeaderService? validateHeaderService)
        {
            HeaderRowService = headerRowService;
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

        protected TModel GetDataFromCell<TModel>(IExcelDataReader reader, List<ExcelColumnModel> columnList,
            out List<ILoggingModel> errorsList, ParsingMethod parsingMethod)
        {
            errorsList = new List<ILoggingModel>();
            var parsingFactory = GetDataFromCellFactory.GetDataFromCell(parsingMethod);
            var result = parsingFactory.MappingData<TModel>(reader, columnList, ref errorsList);

            return result;
        }

        protected string[]? GetHeaderRowInfo(IExcelDataReader reader, int lengthOfHeader)
        {
            if (lengthOfHeader <= 0)
            {
                return null;
            }

            var rowContents = new List<string[]>
            {
                ExcelHelper.ReadEntireRow(reader)
            };

            for (var i = 1; i < lengthOfHeader; i++)
            {
                var canReadNextRow = reader.Read();
                if (!canReadNextRow)
                {
                    return null;
                }

                rowContents.Add(ExcelHelper.ReadEntireRow(reader));
            }

            var traverseList = FlipList(rowContents);
            var result = traverseList.Select(n => n.LastOrDefault(m => !string.IsNullOrWhiteSpace(m)))
                .ToList();

            return result.ToArray()!;
        }

        private List<string?[]> FlipList(List<string[]> input)
        {
            var longestRow = input.Max(n => n.Length);
            var result = new List<string?[]>();
            for (var i = 0; i < longestRow; i++)
            {
                var fragment = new List<string?>();
                foreach (var inputItem in input)
                {
                    fragment.Add(inputItem.Length - 1 < i ? null : inputItem[i]);
                }

                result.Add(fragment.ToArray());
            }

            return result;
        }
    }
}