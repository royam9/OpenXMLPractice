using Microsoft.AspNetCore.Mvc;
using Models;
using Services.Interfaces;
using static Services.WordService;

namespace OpenXMLPractice.Controllers;

[ApiController]
[Route("[controller]")]
public class HomeController : Controller
{
    private readonly IExcelService _excelService;
    private readonly IGeneralService _generalService;

    public HomeController(IExcelService excelService,
        IGeneralService generalService)
    {
        _excelService = excelService;
        _generalService = generalService;
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

        filePath = @"C:\Users\TWJOIN\Desktop\UnknowProject\tryLineChart.xlsx";
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
        string filePath = @"C:\Users\TWJOIN\Desktop\UnknowProject\tryInputChart.docx";

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
        string filePath = @"C:\Users\TWJOIN\Desktop\UnknowProject\tryInputChart.docx";

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
        string filePath = @"C:\Users\TWJOIN\Desktop\UnknowProject\tryInputChart.docx";

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
}
