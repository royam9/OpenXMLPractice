using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Drawing;
using Services.Interfaces;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml;
using Chart = DocumentFormat.OpenXml.Drawing.Charts.Chart;
using Index = DocumentFormat.OpenXml.Drawing.Charts.Index;
using Drawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Formula = DocumentFormat.OpenXml.Drawing.Charts.Formula;
using Values = DocumentFormat.OpenXml.Drawing.Charts.Values;
using Outline = DocumentFormat.OpenXml.Drawing.Outline;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;
using Models;

namespace Services;

/// <summary>
/// Word相關服務
/// </summary>
public class WordService
{
    public class Solution1
    {
        public async Task AddExcelChartToExistingWordDocument(string filePath)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
            {
                // 新增 ChartPart
                MainDocumentPart mainPart = wordDoc.MainDocumentPart;
                ChartPart chartPart = mainPart.AddNewPart<ChartPart>();

                // 設定圖表 ID
                string chartId = "rId" + (mainPart.GetPartsOfType<ChartPart>().Count() + 1).ToString();

                // 設定圖表的資料
                chartPart.ChartSpace = new DocumentFormat.OpenXml.Drawing.Charts.ChartSpace();
                chartPart.ChartSpace.AppendChild(new DocumentFormat.OpenXml.Drawing.Charts.EditingLanguage() { Val = "en-US" });

                var chart = chartPart.ChartSpace.AppendChild(new DocumentFormat.OpenXml.Drawing.Charts.Chart());
                chart.AppendChild(new AutoTitleDeleted() { Val = true });

                // 設定圖表的資料系列
                var plotArea = chart.AppendChild(new PlotArea());
                var layout = plotArea.AppendChild(new Layout());
                var lineChart = plotArea.AppendChild(new LineChart());
                lineChart.AppendChild(new Grouping() { Val = GroupingValues.Standard });
                lineChart.AppendChild(new VaryColors() { Val = false });

                // 資料系列 (Series)
                var series = lineChart.AppendChild(new LineChartSeries());
                series.AppendChild(new Index() { Val = (UInt32Value)0U });
                series.AppendChild(new Order() { Val = (UInt32Value)0U });
                series.AppendChild(new SeriesText(new NumericValue() { Text = "Sample Series" }));

                // X 軸資料
                var categoryAxisData = series.AppendChild(new CategoryAxisData());
                var stringReference = categoryAxisData.AppendChild(new StringReference());
                stringReference.AppendChild(new Formula() { Text = "Sheet1!$A$2:$A$6" });
                var stringCache = stringReference.AppendChild(new StringCache());
                stringCache.AppendChild(new PointCount() { Val = (UInt32Value)5U });

                for (int i = 0; i < 5; i++)
                {
                    var point = stringCache.AppendChild(new StringPoint() { Index = (UInt32Value)(uint)i });
                    point.AppendChild(new NumericValue() { Text = "Category " + (i + 1) });
                }

                // Y 軸資料
                var values = series.AppendChild(new Values());
                var numberReference = values.AppendChild(new NumberReference());
                numberReference.AppendChild(new Formula() { Text = "Sheet1!$B$2:$B$6" });
                var numberCache = numberReference.AppendChild(new NumberingCache());
                numberCache.AppendChild(new PointCount() { Val = (UInt32Value)5U });

                for (int i = 0; i < 5; i++)
                {
                    var point = numberCache.AppendChild(new NumericPoint() { Index = (UInt32Value)(uint)i });
                    point.AppendChild(new NumericValue() { Text = (i * 10).ToString() });
                }

                // 設定軸
                plotArea.AppendChild(new CategoryAxis(new AxisId() { Val = 48650112U }));
                plotArea.AppendChild(new ValueAxis(new AxisId() { Val = 48672768U }));

                lineChart.AppendChild(new AxisId() { Val = 48650112U });
                lineChart.AppendChild(new AxisId() { Val = 48672768U });

                chartPart.ChartSpace.Save();

                // 將圖表添加到 Word 文件中的段落
                AddChartToDocument(mainPart, chartId);
            }
        }

        private static void AddChartToDocument(MainDocumentPart mainPart, string chartId)
        {
            var paragraph = mainPart.Document.Body.AppendChild(new Paragraph());
            var run = paragraph.AppendChild(new Run());

            // 插入圖表
            var drawing = run.AppendChild(new Drawing());
            var inline = drawing.AppendChild(new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline());
            inline.AppendChild(new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent() { Cx = 5486400, Cy = 3200400 });
            inline.AppendChild(new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent()
            {
                LeftEdge = 19050L,
                TopEdge = 0L,
                RightEdge = 9525L,
                BottomEdge = 0L
            });
            inline.AppendChild(new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties() { Id = (UInt32Value)1U, Name = "Chart" });

            var graphic = inline.AppendChild(new DocumentFormat.OpenXml.Drawing.Graphic());
            var graphicData = graphic.AppendChild(new DocumentFormat.OpenXml.Drawing.GraphicData()
            {
                Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"
            });

            graphicData.AppendChild(new DocumentFormat.OpenXml.Drawing.Charts.ChartReference() { Id = chartId });
        }
    }

    public class CSDNSolution
    {
        public async Task<byte[]> AddExcelChartToExistingWordDocument(string filePath)
        {
            using FileStream fileStream = new(filePath, FileMode.Open, FileAccess.Read);
            using MemoryStream memoryStream = new();

            await fileStream.CopyToAsync(memoryStream);

            memoryStream.Position = 0;

            using WordprocessingDocument wordDoc = WordprocessingDocument.Open(memoryStream, true);

            MainDocumentPart mainPart = wordDoc.MainDocumentPart;

            // 添加圖表
            ChartPart chartPart = mainPart.AddNewPart<ChartPart>();
            // 初始化一個ChartSpace
            // ChartSpace：存放所有圖表設定與資料的容器
            // EditingLanguage：定義圖表的編輯語言
            chartPart.ChartSpace = new ChartSpace(new EditingLanguage() { Val = "zh-tw" });
            Chart chart = chartPart.ChartSpace.AppendChild(new Chart());
            // 在繪圖區域加入一條折線
            // PlotArea：定義圖表中「繪圖區域」的內容
            // LineChartSeries：資料系列(Series)，本案即折線圖中的一條折線
            chart.PlotArea = new PlotArea();
            LineChartSeries lineChartSeries = new LineChartSeries();

            uint index = 0;
            // 儲存折線的名稱的變數
            string seriesText = "Series 1";
            // 定義Excel工作表中的範圍(此為Sheet1的A1~B5)
            // Sheet1 工作表名稱
            // ! 分隔工作表及表格描述
            // $ 絕對引用，無論範圍被複製或移動，其位置都不會改變
            // : 應該是 到 的意思
            string rangerefernece = "Sheet1!$A$1:$B$5";

            #region 折線圖數值相關設定
            // 為該折線添加Y軸數值資料容器
            // Values：用於表示圖表中 Y 軸數值資料的容器
            // AppendChild：為該結構添加子節點，返回添加的節點的Reference
            Values values = lineChartSeries.AppendChild(new Values());
            // 將用來指向外部參考的節點添加進Values裡
            // 表示這條折線的數值資料是參考外部
            // NumberReference：用於指定圖表的數值資料來源
            NumberReference numberReference = values
                .AppendChild(new NumberReference());
            // 在NumberReference裡面添加明確定義範圍的子節點
            // Formula：定義數據來源的範圍，藉由字串定義
            numberReference.AppendChild(new Formula(rangerefernece));
            // 添加 NumberingCache 節點，用於內嵌數據
            // (也避免外部數據丟失時圖表變為空白)
            // (數據如果是靜態的(不來自外部)，則不需Formula，直接填充NumberingCache)
            numberReference.AppendChild(new NumberingCache());
            #endregion

            #region 圖表基本資訊設定
            // 為該折線添加索引，用以區分是哪一條折線(資料系列)
            lineChartSeries.AppendChild(new Index() { Val = new UInt32Value(index) });
            // 指定該折線的繪製順序
            lineChartSeries.AppendChild(new Order() { Val = new UInt32Value(index) });
            // SeriesText：定義該資料系列的名稱，也就是在圖表的圖例（Legend）中顯示的文字
            // NumericValue：用於包裹名稱的文字內容
            lineChartSeries.AppendChild(new SeriesText(new NumericValue() { Text = seriesText }));
            #endregion

            #region 樣式設定
            // 將線條設定為無填滿
            // ChartShapeProperties：定義圖表中形狀或線條的樣式屬性
            lineChartSeries.ChartShapeProperties = new ChartShapeProperties();
            lineChartSeries.ChartShapeProperties.AppendChild(new NoFill());
            // 設定線條顏色為黑色
            SolidFill solidFill = new SolidFill();
            RgbColorModelHex rgbColorModelHex = new RgbColorModelHex() { Val = "000000" };
            solidFill.AppendChild(rgbColorModelHex);
            lineChartSeries.ChartShapeProperties.AppendChild(solidFill);
            // DataLabels：控制顯示數據標籤的元素，數據標籤通常在圖表的資料點旁邊，
            //             用來顯示該點的具體數值或其他訊息
            DataLabels dataLabels = lineChartSeries.AppendChild(new DataLabels());
            // 設定資料點旁會顯示具體數值，加入DataLabels設定
            dataLabels.AppendChild(new ShowValue() { Val = true });
            #endregion

            LineChart lineChart = chart.PlotArea.AppendChild(new LineChart());
            lineChart.AppendChild(lineChartSeries);

            // 將折線類型設為折線圖(失敗，沒有ChartType和ChartSubType屬性)            
            // lineChartSeries.ChartType = lineChart.ChartSubType;

            byte[] modifiedDocBytes = memoryStream.ToArray();

            return modifiedDocBytes;
        }
    }

    public class UpdateLineChartExcelValueSolution
    {
        /// <summary>
        /// 通用服務
        /// </summary>
        private readonly IGeneralService _generalService;

        public UpdateLineChartExcelValueSolution(IGeneralService generalService)
        {
            _generalService = generalService;
        }

        public async Task<byte[]> UpdateLineChartExcelValue(string filePath, string cellReference, string cellvalue)
        {
            using FileStream fileStream = new(filePath, FileMode.Open, FileAccess.Read);
            using MemoryStream memoryStream = new();

            await fileStream.CopyToAsync(memoryStream);

            memoryStream.Position = 0;

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(memoryStream, true))
            {

                MainDocumentPart mainPart = wordDoc.MainDocumentPart;

                // 找到文檔中的 ChartPart
                ChartPart chartPart = wordDoc.MainDocumentPart.ChartParts.FirstOrDefault();

                #region 變更內嵌Excel表單
                EmbeddedPackagePart embeddedExcel = chartPart.EmbeddedPackagePart;

                // 取得Excel的Stream
                var excelStream = embeddedExcel.GetStream(FileMode.Open, FileAccess.ReadWrite);

                // 透過SpreadsheetDocument.Open開啟，引數2設為true，可以變更
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(excelStream, true))
                {
                    // 取得WorkbookPart
                    var workbookPart = spreadsheetDocument.WorkbookPart;
                    // 取得第一個WorksheetPart
                    var worksheetPart = workbookPart.WorksheetParts.FirstOrDefault();
                    // 取得第一個SheetData
                    var sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
                    var targetRow = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == _generalService.GetRowIndex(cellReference));
                    var targetCell = targetRow.Elements<Cell>().FirstOrDefault(c => c.CellReference == cellReference);

                    targetCell.CellValue = new CellValue(cellvalue);
                    targetCell.DataType = new EnumValue<CellValues>(CellValues.Number); // 指定數值類型

                    worksheetPart.Worksheet.Save();
                }
                #endregion

                #region 同步變更快取
                var SeriesList = chartPart.ChartSpace.Descendants<LineChartSeries>().ToList();

                // 我要找第一條 第一條就是A2A5 + B2B5
                // 他會是第一個LineChartSeries
                var targetlineChartSeries = chartPart.ChartSpace.Descendants<LineChartSeries>().FirstOrDefault();
                // 找到裡面記載Values(數值資訊)的板塊
                // ※不能直接找LineChartSeries的Descendants<NumericPoint>，除了Values以外還有其他子節點有NumericPoint
                var values = targetlineChartSeries.Elements<Values>().FirstOrDefault();
                // Values裡面的NumerPoint就是對應該條線上的每一個點
                // 這邊找第一個點(B2)
                var targetnumericPoint = values.Descendants<NumericPoint>().FirstOrDefault();
                targetnumericPoint.NumericValue.Text = cellvalue;

                chartPart.ChartSpace.Save();
                #endregion
            }

            //memoryStream.Position = 0;

            return memoryStream.ToArray();
        }

        /// <summary>
        /// 更新圖表
        /// </summary>
        /// <param name="filePath">檔案路徑</param>
        /// <param name="chartTitle">圖表名稱</param>
        /// <param name="param">輸入參數</param>
        /// <returns></returns>
        public async Task<byte[]> UpdateChart(string filePath, string chartTitle, List<List<string>> param)
        {
            using FileStream fileStream = new(filePath, FileMode.Open, FileAccess.Read);
            using MemoryStream memoryStream = new();

            await fileStream.CopyToAsync(memoryStream);

            memoryStream.Position = 0;
            using WordprocessingDocument wordDoc = WordprocessingDocument.Open(memoryStream, true);

            // 尋找目標ChartPart
            ChartPart? targetChartPart = GetChartPart(chartTitle, wordDoc);

            UpdateInnerExcel(param, targetChartPart);

            UpdateCache(param, targetChartPart);

            wordDoc.Save();
            wordDoc.Dispose();

            return memoryStream.ToArray();
        }

        /// <summary>
        /// 取得目標圖表Part
        /// </summary>
        /// <param name="chartTitle">圖表標題</param>
        /// <param name="wordDoc">文件主體</param>
        /// <returns></returns>
        /// <exception cref="Exception">找不到目標圖表</exception>
        private static ChartPart GetChartPart(string chartTitle, WordprocessingDocument wordDoc)
        {
            ChartPart? targetChartPart = null;

            foreach (var chartPart in wordDoc.MainDocumentPart.ChartParts)
            {
                var title = chartPart.ChartSpace.Elements<Chart>().FirstOrDefault()?.Title?.InnerText;

                if (title == chartTitle)
                    targetChartPart = chartPart;
            }

            if (targetChartPart == null)
                throw new Exception(string.Format(ErrorMessages.DataNotFound, "目標圖表"));

            return targetChartPart;
        }

        /// <summary>
        /// 更新內嵌圖表的Excel
        /// </summary>
        /// <param name="param">輸入參數</param>
        /// <param name="targetChartPart">目標圖表Part</param>
        private void UpdateInnerExcel(List<List<string>> param, ChartPart? targetChartPart)
        {
            EmbeddedPackagePart embeddedExcel = targetChartPart.EmbeddedPackagePart;

            // 取得Excel的Stream
            var excelStream = embeddedExcel.GetStream(FileMode.Open, FileAccess.ReadWrite);

            using SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(excelStream, true);
            // 取得WorkbookPart
            var workbookPart = spreadsheetDocument.WorkbookPart;
            // 取得第一個WorksheetPart
            var worksheetPart = workbookPart.WorksheetParts.FirstOrDefault();
            // 取得第一個SheetData
            var sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();

            var rows = sheetData.Elements<Row>().ToList();

            for (int i = 0; i < rows.Count; i++)
            {
                // 第一行是標題，不用填
                if (i == 0)
                    continue;

                Row row = rows[i];
                var cells = row.Elements<Cell>().ToList();

                for (int c = 0; c < cells.Count; c++)
                {
                    Cell cell = cells[c];
                    string? value = param.ElementAtOrDefault(i - 1)?.ElementAtOrDefault(c);

                    // 如果沒有對應項目，將表格設為空
                    if (string.IsNullOrEmpty(value))
                    {
                        cell.CellValue = new CellValue();
                        continue;
                    }

                    // 第一行對應參數位置0的List<string>
                    if (c == 0)
                        cell.CellValue = new CellValue(_generalService.ConvertToExcelDate(value).ToString());
                    else
                        cell.CellValue = new CellValue(value);
                }
            }

            worksheetPart.Worksheet.Save();
        }

        /// <summary>
        /// 更新圖表快取 (快取影響圖表顯示)
        /// </summary>
        /// <param name="param">輸入參數</param>
        /// <param name="targetChartPart">目標圖表Part</param>
        private void UpdateCache(List<List<string>> param, ChartPart? targetChartPart)
        {
            var SeriesList = targetChartPart.ChartSpace.Descendants<LineChartSeries>().ToList();

            for (int i = 0; i < SeriesList.Count; i++)
            {
                var series = SeriesList[i];

                // 一條線(Series)有兩個NumberingCache，一個紀錄X軸數值
                // 一個紀錄X軸對應的Y軸數值
                var numberingCaches = series.Descendants<NumberingCache>().ToList();

                for (int nc = 0; nc < numberingCaches.Count; nc++)
                {
                    var numberingCache = numberingCaches[nc];

                    var numericPoints = numberingCache.Descendants<NumericPoint>().ToList();

                    // 第一個NumberingCache，在Point填入X軸數值(每個List<string>的第一個string)
                    if (nc == 0)
                    {
                        // 將X軸格式改為通用
                        numberingCache.FormatCode.Text = "yyyy-mm-dd";

                        for (int j = 0; j < numericPoints.Count; j++)
                        {
                            // 如果無對應X軸資料，該項目設為空
                            string? value = param.ElementAtOrDefault(j)?.ElementAtOrDefault(0);

                            if (string.IsNullOrEmpty(value))
                            {
                                numericPoints[j].NumericValue.Text = string.Empty;
                                // 使用continue確保無數值的格子被設為空
                                continue;
                            }

                            numericPoints[j].NumericValue.Text = _generalService.ConvertToExcelDate(value).ToString();
                        }
                    }
                    // 第二個NumberingCache，在Point填入X軸對應的Y軸數值
                    else
                    {
                        numberingCache.FormatCode.Text = "0.##";

                        for (int j = 0; j < numericPoints.Count; j++)
                        {
                            string? value = param.ElementAtOrDefault(j)?.ElementAtOrDefault(i + 1);

                            if (string.IsNullOrEmpty(value))
                            {
                                numericPoints[j].NumericValue.Text = string.Empty;
                                continue;
                            }

                            // 每行第一個是x軸 所以i+1
                            numericPoints[j].NumericValue.Text = value;
                        }
                    }
                }
            }

            targetChartPart.ChartSpace.Save();
        }
    }
}