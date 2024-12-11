using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Services.Interfaces;

namespace Services;

/// <summary>
/// Excel相關服務
/// </summary>
public class ExcelService : IExcelService
{
    private readonly IGeneralService _generalService;

    public ExcelService(IGeneralService generalService)
    {
        _generalService = generalService;
    }

    /// <summary>
    /// 輸入值進儲存格
    /// </summary>
    /// <param name="filePath">檔案位置</param>
    /// <param name="sheetName">工作表名稱</param>
    /// <param name="cellReference">儲存格位置</param>
    /// <param name="value">輸入的值</param>
    /// <returns></returns>
    public async Task<byte[]> InsertValueIntoCell(string filePath, string sheetName, string cellReference, string value)
    {
        using MemoryStream memoryStream = new();
        using FileStream fileStream = new(filePath, FileMode.Open, FileAccess.Read);

        await fileStream.CopyToAsync(memoryStream);

        memoryStream.Position = 0;

        // 開啟 Excel 檔案
        using SpreadsheetDocument document = SpreadsheetDocument.Open(memoryStream, true);

        // 獲取 WorkbookPart 和對應的 WorksheetPart
        WorkbookPart workbookPart = document.WorkbookPart;
        Sheet sheet = workbookPart.Workbook.Descendants<Sheet>()
            .FirstOrDefault(s => s.Name == sheetName);

        if (sheet == null)
            throw new ArgumentException($"找不到工作表：{sheetName}");

        WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);

        // 獲取或創建目標儲存格
        Cell cell = GetOrCreateCell(worksheetPart.Worksheet, cellReference);

        // 插入值並更新檔案
        cell.CellValue = new CellValue(value);

        // 根據值指定輸入的儲存格類型
        if (int.TryParse(value, out _))
            cell.DataType = new EnumValue<CellValues>(CellValues.Number); 
        else
            cell.DataType = new EnumValue<CellValues>(CellValues.String);

        worksheetPart.Worksheet.Save();
        document.Dispose();

        return memoryStream.ToArray();
    }

    /// <summary>
    /// 取得該儲存格
    /// </summary>
    /// <param name="worksheet"></param>
    /// <param name="cellReference">儲存格位置</param>
    /// <returns></returns>
    private  Cell GetOrCreateCell(Worksheet worksheet, string cellReference)
    {
        SheetData sheetData = worksheet.GetFirstChild<SheetData>();
        Row row = sheetData.Elements<Row>()
            .FirstOrDefault(r => r.RowIndex == _generalService.GetRowIndex(cellReference));

        if (row == null)
        {
            row = new Row { RowIndex = _generalService.GetRowIndex(cellReference) };
            sheetData.Append(row);
        }

        Cell cell = row.Elements<Cell>()
            .FirstOrDefault(c => c.CellReference == cellReference);

        if (cell == null)
        {
            cell = new Cell { CellReference = cellReference };
            row.Append(cell);
        }

        return cell;
    }
}
