using ExcelReaderMapper.Common;
using ExcelReaderMapper.Testing.Models;

namespace ExcelReaderMapper.Testing.Tests;

public class ContentUnitTest : BaseUnitTest
{
    [Fact]
    public void SimpleModelTest_CommonUse()
    {
        var fileContent = File.ReadAllBytes("EmbeddedFile/SimpleModel.xlsx");
        var result = _excelMapperService.GetDataFromExcel<SimpleModel>(fileContent);

        var firstSheet = result.First();
        Assert.True(firstSheet.IsError);
        Assert.False(firstSheet.RowResults[0].IsError);
        Assert.True(firstSheet.RowResults[1].IsError);
        Assert.Contains(firstSheet.RowResults[1].Errors, n => n.Code == ErrorCodeEnum.InvalidField);
        Assert.True(firstSheet.RowResults[2].IsError);
        Assert.Contains(firstSheet.RowResults[2].Errors, n => n.Code == ErrorCodeEnum.MustInDateFormat);
    }

    [Fact]
    public void TwoLayerModelTest_CommonUse()
    {
        var fileContent = File.ReadAllBytes("EmbeddedFile/TwoLayerModel.xlsx");
        var result = _excelMapperService.GetDataFromExcel<SimpleModel>(fileContent, lengthOfHeader: 2);

        var firstSheet = result.First();
        Assert.True(firstSheet.IsError);
        Assert.False(firstSheet.RowResults[0].IsError);
        Assert.True(firstSheet.RowResults[1].IsError);
        Assert.Contains(firstSheet.RowResults[1].Errors, n => n.Code == ErrorCodeEnum.InvalidField);
        Assert.True(firstSheet.RowResults[2].IsError);
        Assert.Contains(firstSheet.RowResults[2].Errors, n => n.Code == ErrorCodeEnum.MustInDateFormat);
    }

    [Fact]
    public void SimpleModelTest_TwoLayer()
    {
        var fileContent = File.ReadAllBytes("EmbeddedFile/SimpleModel.xlsx");
        var result = _excelMapperService.GetDataFromExcel<ModelWithIndex>(fileContent);

        var firstSheet = result.First();
        Assert.True(firstSheet.IsError);
        Assert.False(firstSheet.RowResults[0].IsError);
        Assert.True(firstSheet.RowResults[1].IsError);
        Assert.Contains(firstSheet.RowResults[1].Errors, n => n.Code == ErrorCodeEnum.InvalidField);
        Assert.True(firstSheet.RowResults[2].IsError);
        Assert.Contains(firstSheet.RowResults[2].Errors, n => n.Code == ErrorCodeEnum.MustInDateFormat);
    }

    [Fact]
    public void TwoLayerModelTest_TwoLayer()
    {
        var fileContent = File.ReadAllBytes("EmbeddedFile/TwoLayerModel.xlsx");
        var result = _excelMapperService.GetDataFromExcel<ModelWithIndex>(fileContent, lengthOfHeader: 2);

        var firstSheet = result.First();
        Assert.True(firstSheet.IsError);
        Assert.False(firstSheet.RowResults[0].IsError);
        Assert.True(firstSheet.RowResults[1].IsError);
        Assert.Contains(firstSheet.RowResults[1].Errors, n => n.Code == ErrorCodeEnum.InvalidField);
        Assert.True(firstSheet.RowResults[2].IsError);
        Assert.Contains(firstSheet.RowResults[2].Errors, n => n.Code == ErrorCodeEnum.MustInDateFormat);
    }
}