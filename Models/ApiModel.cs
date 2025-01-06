using Microsoft.AspNetCore.Http;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

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

    public class GetInnerXMLRequestModel
    {
        [Required]
        public IFormFile File { get; set; } = null!;
    }
    #endregion

    #region Response

    #endregion
}
