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
}
