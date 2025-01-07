using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;
using Models;
using Models.AntioxidantReportModel;
using Services;
using Services.Interfaces;
using Services.Interfaces.AntioxidantReport;
using System.IO;
using static Services.WordService;

namespace OpenXMLPractice.Controllers;

[ApiController]
[Route("[controller]")]
public class HomeController : Controller
{
    private readonly IExcelService _excelService;
    private readonly IGeneralService _generalService;
    private readonly IAntioxidantReportService _antioxidantReportService;

    public HomeController(IExcelService excelService,
        IGeneralService generalService,
        IAntioxidantReportService antioxidantReportService)
    {
        _excelService = excelService;
        _generalService = generalService;
        _antioxidantReportService = antioxidantReportService;
    }

    /// <summary>
    /// 在Excel指定儲存格輸入值
    /// </summary>
    /// <param name="cellReference">儲存格位置</param>
    /// <param name="value">值</param>
    /// <returns></returns>
    [HttpPost]
    [Route("InputExcel")]
    public async Task<IActionResult> InputExcel([FromForm] string cellReference = "B4", [FromForm] string value = "123")
    {
        string filePath, sheetName;

        filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "Template", "tryLineChart.xlsx");
        sheetName = "Table01";

        var excelbyte = await _excelService.InsertValueIntoCell(
            filePath,
            sheetName,
            cellReference,
            value);

        return File(excelbyte, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "example.xlsx");
    }

    /// <summary>
    /// 在Word插入折線圖
    /// </summary>
    /// <returns></returns>
    [HttpPost]
    [Route("InputChartAtWord")]
    public async Task<IActionResult> InputChartAtWord()
    {
        string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "Template", "tryInputChart.docx");

        var wordTool = new CSDNSolution();

        var result = await wordTool.AddExcelChartToExistingWordDocument(filePath);

        return File(result, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "example.docx");
    }

    /// <summary>
    /// 在Word更新折線圖的數值
    /// </summary>
    /// <param name="cellReference">儲存格位置</param>
    /// <param name="cellValue">欲輸入的值</param>
    /// <returns></returns>
    [HttpPost]
    [Route("UpdateLineChartExcelValue")]
    public async Task<IActionResult> UpdateLineChartExcelValue([FromForm] string cellReference = "B2", [FromForm] string cellValue = "100")
    {
        string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "Template", "tryInputChart.docx");

        var updateTool = new UpdateLineChartExcelValueSolution(_generalService);

        var result = await updateTool.UpdateLineChartExcelValue(filePath, cellReference, cellValue);

        return File(result, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "example.docx");
    }

    /// <summary>
    /// 在Word更新整個折線圖的數值
    /// </summary>
    /// <param name="param">輸入參數</param>
    /// <returns></returns>
    /// <remarks>InputData的一個List string 記載一條X軸的資料 <br />
    /// List string 的第一個項目固定為日期格式(yyyy/mm/dd) <br />
    /// 其他List string裡的項目限定為數字</remarks>
    [HttpPost]
    [Route("UpdateChartExcelValue")]
    public async Task<IActionResult> UpdateChartExcelValue([FromBody] UpdateChartExcelValueRequestModel param)
    {
        string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "Template", "tryInputChart.docx");

        if (string.IsNullOrEmpty(param.ChartTitle))
            param.ChartTitle = "特性成長趨勢圖";

        if (param.InputData == null || param.InputData.Count == 0)
        {
            param.InputData = [
                ["2024/10/10", "10", "20", "30", "40"],
                ["2024/10/11", "15", "25", "35", "45"],
                ["2024/10/12", "16", "26", "37", "46"],
                ["2024/10/13", "17", "27", "38", "47"]
            ];
        }

        var updateTool = new UpdateLineChartExcelValueSolution(_generalService);

        var result = await updateTool.UpdateChart(filePath, param.ChartTitle, param.InputData);

        return File(result, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "example.docx");
    }

    /// <summary>
    /// 在Word新增浮水印
    /// </summary>
    /// <returns></returns>
    [HttpPost]
    [Route("InsertWatermarkInWord")]
    public async Task<IActionResult> InsertWatermarkInWord()
    {
        var tool = new WordWatermarkService();

        string docPath = @"C:\Users\TWJOIN\Desktop\安寶\報告輸出模板\Hi.docx";
        string picPath = @"C:\Users\TWJOIN\Desktop\安寶\報告輸出模板\安寶報告章.驗收章.png";

        var result = await tool.InsertWatermark2(docPath, picPath);

        return File(result, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "outputTest.docx");
    }

    [HttpPost]
    [Route("GetInnerXML")]
    public async Task<IActionResult> GetInnerXML([FromForm] GetInnerXMLRequestModel param)
    {
        using MemoryStream memoryStream = new();
        await param.File.CopyToAsync(memoryStream);
        using WordprocessingDocument package = WordprocessingDocument.Open(memoryStream, true);
        MainDocumentPart? mainDocPart = package.MainDocumentPart;
        Body? body = mainDocPart?.Document.Body;
        return Ok();
    }

    [HttpPost]
    [Route("GenerateAntioxidantReport")]
    public async Task<IActionResult> GenerateAntioxidantReport()
    {
        List<AntioxidantReportTransformerBaseInfoModel> transformerInfo = new();
        AntioxidantReportTransformerBaseInfoModel model = new()
        {
            Number = "1",
            TransformerName = "我是測試變壓器",
            TransformerSerialNumber = "我是製造號碼",
            SamplingOilTemperature = "100°C",
            AntioxidantContent = "100"
        };
        transformerInfo.Add(model);
        transformerInfo.Add(model);

        var theModel = new AntioxidantReportModel
        {
            SampleProvider = "我是委託單位",
            SampleProviderAddress = "我是委託單位地址",
            ExperimentSerialNumber = "我是試驗編號",
            IssueDate = "我是發行日期",
            SampleCount = "我是件數",
            SampleDate = "我是取樣日期",
            ExperimentDate = "我是試驗時間",
            Sampler = "我是測驗人",
            TransformerInfo = transformerInfo
        };

        var result = await _antioxidantReportService.GenerateAntioxidantReport(theModel);
        return File(result, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "AntioxidantReportTest.docx");
    }
}
