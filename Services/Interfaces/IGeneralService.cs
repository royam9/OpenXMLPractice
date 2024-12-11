namespace Services.Interfaces;

/// <summary>
/// 通用服務
/// </summary>
public interface IGeneralService
{
    /// <summary>
    /// 取得 儲存格位置 中的行號(數字部分)
    /// </summary>
    /// <param name="cellReference">儲存格位置</param>
    /// <returns>無符號整數 (代表不會是負的)</returns>
    uint GetRowIndex(string cellReference);
    /// <summary>
    /// 將日期字串轉換成數字格式
    /// </summary>
    /// <param name="dateString">日期字串</param>
    /// <returns></returns>
    /// <exception cref="FormatException"></exception>
    double ConvertToExcelDate(string dateString);
}
