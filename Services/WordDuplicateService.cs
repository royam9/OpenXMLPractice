using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Models.AntioxidantReportModel;

namespace Services;

public static class WordDuplicateService
{    
    //var startParagraph = bookmarkStarts.FirstOrDefault(b => b.Name == "StartParagraph").Ancestors<Paragraph>().FirstOrDefault();
    //var endParagraph = bookmarkStarts.FirstOrDefault(b => b.Name == "EndParagraph").Ancestors<Paragraph>().FirstOrDefault();
    //var body = startParagraph.Ancestors<Body>().FirstOrDefault();
    //var elementsToCopy = body.Elements()
    //          .SkipWhile(e => e != startParagraph) // 從起始段落開始
    //          .TakeWhile(e => e != endParagraph)  // 直到結束段落
    //          .ToList();

    //elementsToCopy.Add(endParagraph);      // 包含結束段落

    //private static void HandleBaseInfo(BookmarkStart bookmarkStart, IEnumerable<AntioxidantReportTransformerBaseInfoModel> transformerListInfo,
    //    List<OpenXmlElement> elementsToCopy)
    //{
    //    var transformerListInfoList = transformerListInfo as List<AntioxidantReportTransformerBaseInfoModel> ?? transformerListInfo.ToList();
    //    Table nowTable = bookmarkStart.Ancestors<Table>().FirstOrDefault();
    //    //BookmarkStart nowbookmarkStart = bookmarkStart;

    //    // BaseInfo如果超過15個
    //    while (transformerListInfoList.Count() > 15)
    //    {
    //        // 複製目前的Table
    //        Table duplicateTable = bookmarkStart.Ancestors<Table>().FirstOrDefault().CloneNode(true) as Table;

    //        // 然後找到裡面的bookmarkStart
    //        BookmarkStart insideBookmarkStart = duplicateTable.Descendants<BookmarkStart>().FirstOrDefault(d => d.Name == bookmarkStart.Name);

    //        // 整理要給function的model
    //        List<AntioxidantReportTransformerBaseInfoModel> insertBaseInfo = transformerListInfoList.Take(15).ToList();

    //        // 進行一樣的事
    //        InsertTransformerInfo(insideBookmarkStart, insertBaseInfo);

    //        // 把原本的bookmarkStart刪掉
    //        //nowbookmarkStart.Remove();

    //        // 新的書籤 = 裡面的這個
    //        //nowbookmarkStart = insideBookmarkStart;

    //        // 把上面那段，複製的Table，和下面的分節符號合起來
    //        var trueElementsToCopyCollection = elementsToCopy.ToList();
    //        trueElementsToCopyCollection.Add(duplicateTable);

    //        // 加在原本的Table後面
    //        foreach (var item in trueElementsToCopyCollection)
    //        {
    //            var openxmlElement = item.CloneNode(true);
    //            nowTable.InsertAfterSelf(openxmlElement);
    //        }

    //        // 把填入過的Info刪掉
    //        transformerListInfoList.RemoveAll(t => insertBaseInfo.Contains(t));

    //        // 現在Table = 加入後的Table
    //        // nowTable = duplicateTable;
    //    }

    //    if (transformerListInfoList.Count > 0)
    //    {
    //        InsertTransformerInfo(bookmarkStart, transformerListInfoList);
    //    }
    //}
}
