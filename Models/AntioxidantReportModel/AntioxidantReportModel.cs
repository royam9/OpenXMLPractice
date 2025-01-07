namespace Models.AntioxidantReportModel;

/// <summary>
/// 抗氧化劑試驗報告 Model
/// </summary>
public class AntioxidantReportModel
{
    /// <summary>
    /// 委託單位
    /// </summary>
    public required string SampleProvider { get; set; }
    /// <summary>
    /// 委託單位地址
    /// </summary>
    public string? SampleProviderAddress { get; set; }
    /// <summary>
    /// 試驗編號
    /// </summary>
    public string? ExperimentSerialNumber { get; set; }
    /// <summary>
    /// 發行日期 (yyyy年MM月dd日)
    /// </summary>
    public string? IssueDate { get; set; }
    /// <summary>
    /// 樣品數量
    /// </summary>
    public string? SampleCount { get; set; }
    /// <summary>
    /// 取樣日期 (yyyy/MM/dd~yyyy/MM/dd)
    /// </summary>
    public string? SampleDate { get; set; }
    /// <summary>
    /// 試驗日期 (yyyy/MM/dd)
    /// </summary>
    public string? ExperimentDate { get; set; }
    /// <summary>
    /// 取樣人員
    /// </summary>
    public string? Sampler { get; set; }
    /// <summary>
    /// 變壓器
    /// </summary>
    public required List<AntioxidantReportTransformerBaseInfoModel> TransformerInfo { get; set; }
}

/// <summary>
/// 抗氧化劑試驗報告變壓器基本資訊 Model
/// </summary>
public class AntioxidantReportTransformerBaseInfoModel
{
    /// <summary>
    /// 件次
    /// </summary>
    public string? Number { get; set; }
    /// <summary>
    /// 變壓器名稱
    /// </summary>
    public string? TransformerName { get; set; }
    /// <summary>
    /// 製造號碼
    /// </summary>
    public string? TransformerSerialNumber { get; set; }
    /// <summary>
    /// 取樣油溫 (單位:°C)
    /// </summary>
    public string? SamplingOilTemperature { get; set; }
    /// <summary>
    /// 抗氧化劑含量 (%)
    /// </summary>
    public string? AntioxidantContent { get; set; }
}
