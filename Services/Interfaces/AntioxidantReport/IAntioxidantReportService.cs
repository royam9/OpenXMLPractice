using Models.AntioxidantReportModel;

namespace Services.Interfaces.AntioxidantReport;

/// <summary>
/// 抗氧化劑試驗報告相關服務
/// </summary>
public interface IAntioxidantReportService
{
    /// <summary>
    /// 生成抗氧化劑試驗報告
    /// </summary>
    /// <param name="param">輸入參數</param>
    /// <returns>試驗報告數據</returns>
    Task<byte[]> GenerateAntioxidantReport(AntioxidantReportModel param);
}
