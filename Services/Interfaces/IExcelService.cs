
namespace Services.Interfaces;
public interface IExcelService
{
    /// <summary>
    /// 輸入值進儲存格
    /// </summary>
    /// <param name="filePath">檔案位置</param>
    /// <param name="sheetName">工作表名稱</param>
    /// <param name="cellReference">儲存格位置</param>
    /// <param name="value">輸入的值</param>
    /// <returns></returns>
    Task<byte[]> InsertValueIntoCell(string filePath, string sheetName, string cellReference, string value);
}
