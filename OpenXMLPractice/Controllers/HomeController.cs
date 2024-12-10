using Microsoft.AspNetCore.Mvc;
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
    /// <param name="cellReference"></param>
    /// <param name="cellValue"></param>
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
}
