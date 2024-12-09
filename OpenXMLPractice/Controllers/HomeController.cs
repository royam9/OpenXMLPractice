using Microsoft.AspNetCore.Mvc;
using Services.Interfaces;
using static Services.WordService;

namespace OpenXMLPractice.Controllers;

[ApiController]
[Route("[controller]")]
public class HomeController : Controller
{
    private readonly IExcelService _excelService;
    private readonly IWordService _wordService;

    public HomeController(IExcelService excelService, 
        IWordService wordService)
    {
        _excelService = excelService;
        _wordService = wordService;
    }

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

    [HttpPost]
    [Route("InputChartAtWord")]
    public async Task<IActionResult> InputChartAtWord()
    {
        string filePath = @"C:\Users\TWJOIN\Desktop\UnknowProject\tryInputChart.docx";

        var result = await _wordService.AddExcelChartToExistingWordDocument(filePath);

        return File(result, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "example.docx");
    }

    [HttpPost]
    [Route("UpdateLineChartExcelValue")]
    public async Task<IActionResult> UpdateLineChartExcelValue()
    {
        string filePath = @"C:\Users\TWJOIN\Desktop\UnknowProject\tryInputChart.docx";

        var updateTool = new UpdateLineChartExcelValueSolution();

        await updateTool.UpdateLineChartExcelValue(filePath);

        return Ok();
    }
}
