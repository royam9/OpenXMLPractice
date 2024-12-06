using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml;
using System.IO;
using Services.Interfaces;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Chart = DocumentFormat.OpenXml.Drawing.Charts.Chart;
using Index = DocumentFormat.OpenXml.Drawing.Charts.Index;
using Drawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Formula = DocumentFormat.OpenXml.Drawing.Charts.Formula;
using Values = DocumentFormat.OpenXml.Drawing.Charts.Values;

namespace Services;

/// <summary>
/// Word相關服務
/// </summary>
public class WordService 
{
    public class GPTSolution : IWordService
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

    public class CSDNSolution : IWordService
    {
        public async Task AddExcelChartToExistingWordDocument(string filePath)
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
            chartPart.ChartSpace = new ChartSpace(new EditingLanguage() { Val = "zh-tw"});
            Chart chart = chartPart.ChartSpace.AppendChild(new Chart());
        }
    }
}
