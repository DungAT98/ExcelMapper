using System;
using ExcelReaderMapper.Common;
using ExcelReaderMapper.Infrastructure;

namespace ExcelReaderMapper.Service
{
    internal static class GetDataFromCellFactory
    {
        internal static IRowMappingData GetDataFromCell(ParsingMethod parsingMethod)
        {
            switch (parsingMethod)
            {
                case ParsingMethod.Reflection:
                {
                    var result = new RowMappingReflection();
                    return result;
                }

                case ParsingMethod.NormalCase:
                {
                    var result = new RowMappingParsing();
                    return result;
                }

                default:
                    throw new Exception("Not implemented");
            }
        }
    }
}