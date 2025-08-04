using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Http;

namespace Services;

/// <summary>
/// Word 驗證 Service
/// </summary>
public static class WordValidateService
{
    /// <summary>
    /// 驗證 Word
    /// </summary>
    /// <param name="formFile">Word formFile</param>
    public static async Task ValidateWord(IFormFile formFile)
    {
        MemoryStream stream = new MemoryStream();
        await formFile.CopyToAsync(stream);
        stream.Position = 0;

        ValidateWord(stream);

        stream.Position = 0;

        #region 測試區域
        //using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(stream, true))
        //{
        //    var tableLooks = wordDoc.MainDocumentPart.Document.Body.Descendants<TableLook>().ToList();

        //    foreach (var item in tableLooks)
        //    {
        //        item.FirstRow = new OnOffValue(OnOffOnlyValues.On);
        //        item.LastRow = new OnOffValue(OnOffOnlyValues.On);
        //        item.FirstColumn = new OnOffValue(OnOffOnlyValues.On);
        //        item.LastColumn = new OnOffValue(OnOffOnlyValues.On);
        //        item.NoHorizontalBand = new OnOffValue(OnOffOnlyValues.On);
        //        item.NoVerticalBand = new OnOffValue(OnOffOnlyValues.On);
        //    }
        //}

        //stream.Position = 0;

        //await File.WriteAllBytesAsync(@$"C:\Users\TWJOIN\Desktop\安寶\報告產出文件\測試檔案{DateTime.Now:MM-dd-mm-ss}.docx", stream.ToArray());
        #endregion

        stream.Dispose();
    }

    /// <summary>
    /// 驗證 Word
    /// </summary>
    /// <param name="wordDataStream">Word Stream</param>
    public static void ValidateWord(Stream wordDataStream)
    {
        using WordprocessingDocument wordDoc = WordprocessingDocument.Open(wordDataStream, true);

        OpenXmlValidator validator = new OpenXmlValidator();

        foreach (ValidationErrorInfo error in validator.Validate(wordDoc))
        {
            Console.WriteLine($"ErrorId: {error.Id}");
            Console.WriteLine($"ErrorType: {error.ErrorType}");
            Console.WriteLine($"ErrorRelatedNode: {error.RelatedNode}");
            Console.WriteLine($"ErrorNode: {error.Node}");
            Console.WriteLine($"ErrorPath: {error.Path?.XPath ?? "Null"}");
            Console.WriteLine($"Error: {error.Description}");
            Console.WriteLine($"----------------------------------");
        }

        Console.WriteLine("///End///");
        Console.WriteLine("///End///");
        Console.WriteLine("///End///");
    }
}
