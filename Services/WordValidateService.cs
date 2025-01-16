using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Microsoft.AspNetCore.Http;
using System.ComponentModel.DataAnnotations;

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
            Console.WriteLine($"ErrorPath: {error.Path}");
            Console.WriteLine($"Error: {error.Description}");
            Console.WriteLine($"----------------------------------");
        }
    }
}
