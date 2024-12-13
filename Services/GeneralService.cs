using Services.Interfaces;

namespace Services;

/// <summary>
/// 通用服務
/// </summary>
public class GeneralService : IGeneralService
{
    /// <summary>
    /// 取得 儲存格位置 中的行號(數字部分)
    /// </summary>
    /// <param name="cellReference">儲存格位置</param>
    /// <returns>無符號整數 (代表不會是負的)</returns>
    public uint GetRowIndex(string cellReference)
    {
        string rowPart = new(cellReference.Where(char.IsDigit).ToArray());
        return uint.Parse(rowPart);
    }

    /// <summary>
    /// 將日期字串轉換成數字格式
    /// </summary>
    /// <param name="dateString">日期字串</param>
    /// <returns></returns>
    /// <exception cref="FormatException"></exception>
    public double ConvertToExcelDate(string dateString)
    {
        if (DateTime.TryParse(dateString, out DateTime date))
        {
            // Excel 日期從 1900/1/1 開始計算，DateTime 與 Excel 的起始點對齊
            return (date - new DateTime(1899, 12, 30)).TotalDays;
        }

        throw new FormatException($"無法將日期 '{dateString}' 轉換為 Excel 日期格式");
    }
}
