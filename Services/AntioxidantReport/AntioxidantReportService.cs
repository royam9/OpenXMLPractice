using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Models.AntioxidantReportModel;
using Services.Interfaces.AntioxidantReport;

namespace Services.AntioxidantReport;

/// <summary>
/// 抗氧化劑試驗報告相關服務
/// </summary>
public class AntioxidantReportService : IAntioxidantReportService
{
    public AntioxidantReportService()
    {
        TemplePath = @"C:\Users\TWJOIN\Desktop\安寶\報告輸出模板\處理過的\其他試驗\抗氧化劑-不診斷報告模板.docx";
        GeneralRunProperties = new
            (
            new RunFonts { EastAsia = "標楷體", ComplexScript = "Times New Roman" },
            new FontSize { Val = "24" }
            );
    }

    /// <summary>
    /// 模板路徑
    /// </summary>
    private string TemplePath { get; set; }
    /// <summary>
    /// 通用樣式
    /// </summary>
    private RunProperties GeneralRunProperties { get; set; }

    /// <summary>
    /// 生成抗氧化劑試驗報告
    /// </summary>
    /// <param name="param">輸入參數</param>
    /// <returns>試驗報告數據</returns>
    public async Task<byte[]> GenerateAntioxidantReport(AntioxidantReportModel param)
    {
        using FileStream fileStream = new(TemplePath, FileMode.Open);
        using MemoryStream memoryStream = new();
        await fileStream.CopyToAsync(memoryStream);
        memoryStream.Position = 0;

        using WordprocessingDocument wordDoc = WordprocessingDocument.Open(memoryStream, true);
        InsertWordValue(wordDoc, param);
        wordDoc.Dispose();

        return memoryStream.ToArray();
    }

    /// <summary>
    /// 新增資料進Word
    /// </summary>
    /// <param name="wordDoc">文件主體</param>
    /// <param name="param">輸入參數</param>
    private void InsertWordValue(WordprocessingDocument wordDoc, AntioxidantReportModel param)
    {
        MainDocumentPart? mainDocumentPart = wordDoc.MainDocumentPart ??
            throw new Exception("文件無內容");

        List<BookmarkStart>? bookmarkStarts = mainDocumentPart.RootElement?
            .Descendants<BookmarkStart>()
            .ToList() ??
                throw new Exception("找不到RootElement");

        InsertValueByBookmark(param, bookmarkStarts);
    }

    /// <summary>
    /// 根據書籤插入資料
    /// </summary>
    /// <param name="param">輸入參數</param>
    /// <param name="bookmarkStarts">全部書籤列表</param>
    private void InsertValueByBookmark(AntioxidantReportModel param, List<BookmarkStart> bookmarkStarts)
    {
        foreach (var bookmarkStart in bookmarkStarts)
        {
            string? bookmarkStartName = bookmarkStart.Name?.ToString();
            Run run = new();
            Text text;
            var runProperties = GeneralRunProperties.CloneNode(true);

            switch (bookmarkStartName)
            {
                case "ExperimentDate1":
                    text = new Text(param.ExperimentDate);
                    break;
                case string _ when bookmarkStartName.Contains("ExperimentSerialNumber"):
                    text = new Text(param.ExperimentSerialNumber);
                    break;
                case "IssueDate1":
                    text = new Text(param.IssueDate);
                    break;
                case "SampleCount1":
                    text = new Text(param.SampleCount);
                    break;
                case "SampleDate1":
                    text = new Text(param.SampleDate);
                    break;
                case string _ when bookmarkStartName.Contains("SampleProvider"):
                    text = new Text(param.SampleProvider);
                    break;
                case "SampleProviderAddress1":
                    text = new Text(param.SampleProviderAddress);
                    break;
                case "Sampler1":
                    text = new Text(param.Sampler);
                    break;
                case "TransformerBaseInfo":
                    InsertTransformerInfo(bookmarkStart, param.TransformerInfo);
                    continue;
                default:
                    continue;
            }

            run.AppendChild(runProperties);
            run.AppendChild(text);
            bookmarkStart.InsertAfterSelf(run);
        }
    }

    /// <summary>
    /// 新增變壓器基礎資料
    /// </summary>
    /// <param name="transformerBaseInfoBookmark">變壓器基礎資料第一格書籤</param>
    /// <param name="transformerListInfo">變壓器基礎資料資訊</param>
    private void InsertTransformerInfo(BookmarkStart transformerBaseInfoBookmark, List<AntioxidantReportTransformerBaseInfoModel> transformerListInfo)
    {
        // 取得 填入變壓器資料表格的第一行
        TableRow? tableRow = transformerBaseInfoBookmark?.Ancestors<TableRow>().FirstOrDefault() ??
            throw new Exception("找不到目標Row");

        // 可輸入變壓器資料的所有行缺一
        IEnumerable<TableRow> transformerInfoTableRows = tableRow.ElementsAfter()
            .OfType<TableRow>()
            .SkipLast(1);

        // 加入書籤那行
        transformerInfoTableRows = transformerInfoTableRows.Prepend(tableRow);

        Queue<TableRow> tableRowsQueue = new(transformerInfoTableRows);

        foreach (var info in transformerListInfo)
        {
            TableRow row = tableRowsQueue.Dequeue();
            UpdateRow(row, info);
        }

        DeleteBlankTableRow(tableRowsQueue);
    }

    /// <summary>
    /// 更新目標資料行
    /// </summary>
    /// <param name="tableRow">目標資料行</param>
    /// <param name="param">變壓器資訊</param>
    private void UpdateRow(TableRow tableRow, AntioxidantReportTransformerBaseInfoModel param)
    {
        // 若模板無預設資料可以省略
        DeleteDefaultRunsInRow(tableRow);

        InsertRowValue(tableRow, param);
    }

    /// <summary>
    /// 刪除該資料行預設資料
    /// </summary>
    /// <param name="tableRow">目標資料行</param>
    private static void DeleteDefaultRunsInRow(TableRow tableRow)
    {
        // 將裡面的Run都刪掉
        List<Run>? runsInRow = tableRow.Descendants<Run>().ToList();

        if (runsInRow?.Count > 0)
        {
            foreach (var run in runsInRow)
            {
                run.Remove();
            }
        }
    }

    /// <summary>
    /// 填入值進該資料行
    /// </summary>
    /// <param name="tableRow">目標資料行</param>
    private void InsertRowValue(TableRow tableRow, AntioxidantReportTransformerBaseInfoModel param)
    {
        // 依序在Paragraph裡面填入新Run
        List<Paragraph> paragraphsInRow = tableRow.Descendants<Paragraph>().ToList();

        for (int i = 0; i < paragraphsInRow.Count; i++)
        {
            Paragraph? paragraph = paragraphsInRow[i];
            Run newRun = new();
            var newRunProperties = GeneralRunProperties.CloneNode(true);
            newRun.Append(newRunProperties);

            switch (i)
            {
                case 0:
                    newRun.AddChild(new Text(param.Number));
                    break;
                case 1:
                    newRun.AddChild(new Text(param.TransformerName));
                    break;
                case 2:
                    newRun.AddChild(new Text(param.TransformerSerialNumber));
                    break;
                case 3:
                    newRun.AddChild(new Text(param.SamplingOilTemperature));
                    break;
                case 4:
                    newRun.AddChild(new Text(param.AntioxidantContent));
                    break;
            }

            paragraph.AddChild(newRun);
        }
    }

    /// <summary>
    /// 刪除空白Row
    /// </summary>
    /// <param name="tableRows">空白的資料行集合</param>
    /// <remarks>有排除最後一行</remarks>
    private static void DeleteBlankTableRow(Queue<TableRow> tableRows)
    {
        while (tableRows.Count > 0)
        {
            TableRow row = tableRows.Dequeue();
            row.Remove();
        }
    }
}
