using System.ComponentModel;

namespace Models
{
    #region Request
    public class UpdateChartExcelValueRequestModel
    {
        [DefaultValue(null)]
        public string? ChartTitle { get; set; }
        [DefaultValue(null)]
        public List<List<string>>? InputData { get; set; }
    }
    #endregion

    #region Response

    #endregion
}
