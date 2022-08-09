using System.Collections.Generic;
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
    }
}