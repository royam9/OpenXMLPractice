using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;
using System.Diagnostics;
using DocumentFormat.OpenXml.Vml;

namespace Services
{
    public class WordWatermarkService
    {
        #region 原版 將圖片先轉換成base64再填入目標元素
        /// <summary>
        /// 圖片base64字串
        /// </summary>
        private string imagePartData = "";

        /// <summary>
        /// 加入浮水印
        /// </summary>
        /// <param name="docPath">模板文件路徑</param>
        /// <param name="picPath">浮水印圖片路徑</param>
        /// <returns>文件數據</returns>
        public byte[] InsertWatermark1(string docPath, string picPath)
        {
            using FileStream fileStream = new(docPath, FileMode.Open);
            using MemoryStream memoryStream = new();
            fileStream.CopyTo(memoryStream);
            memoryStream.Position = 0;

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(memoryStream, true))
            {
                InsertCustomWatermark1(wordDoc, picPath);
            }

            return memoryStream.ToArray();
        }

        /// <summary>
        /// 加入客製浮水印 - 有將圖片轉成base64的版本
        /// </summary>
        /// <param name="wordDoc">文件主體</param>
        /// <param name="picPath">圖片路徑</param>
        /// <exception cref="Exception"></exception>
        private void InsertCustomWatermark1(WordprocessingDocument wordDoc, string picPath)
        {
            // 取得圖片的base64字串
            SetWaterMarkPicture(picPath);

            MainDocumentPart? mainDocumentPart = wordDoc.MainDocumentPart;

            if (mainDocumentPart != null)
            {
                // 刪除原本的HeaderParts
                mainDocumentPart.DeleteParts(mainDocumentPart.HeaderParts);

                // 創建新的HeaderParts
                HeaderPart headPart = mainDocumentPart.AddNewPart<HeaderPart>();

                // 在HeaderParts裡設定Header，繪製浮水印區塊
                GenerateHeaderPartContent(headPart);

                // 取得該part的relationshipId
                string rId = mainDocumentPart.GetIdOfPart(headPart);

                // 創建For Image的Part
                ImagePart image = headPart.AddNewPart<ImagePart>("image/png", "rId999");

                // 輸入Image的數據
                GenerateImagePartContent(image);

                SetHeaderPartReference(mainDocumentPart, rId);
            }
            else
            {
                throw new Exception("文件無內容");
            }
        }

        /// <summary>
        /// 讀取浮水印圖片
        /// </summary>
        /// <param name="file">浮水印圖片路徑</param>
        public void SetWaterMarkPicture(string file)
        {
            FileStream inFile;
            byte[] byteArray;
            try
            {
                inFile = new FileStream(file, FileMode.Open, FileAccess.Read);
                byteArray = new byte[inFile.Length];
                long byteRead = inFile.Read(byteArray, 0, (int)inFile.Length);
                inFile.Close();
                imagePartData = Convert.ToBase64String(byteArray, 0, byteArray.Length);
            }
            catch (Exception ex)
            {
                Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// 填入圖片Part內容
        /// </summary>
        /// <param name="imagePart">ImagePart</param>
        /// <remarks>flieStream > bytes > base64 > bytes > memoryStream</remarks>
        private void GenerateImagePartContent(ImagePart imagePart)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePartData);
            imagePart.FeedData(data);
            data.Close();
        }

        /// <summary>
        /// base64字串讀取進Stream
        /// </summary>
        /// <param name="base64String">base64字串</param>
        /// <returns>Stream</returns>
        private System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }
        #endregion

        #region 新版 直接將讀取圖片數據填入目標元素
        /// <summary>
        /// 加入浮水印
        /// </summary>
        /// <param name="docPath">模板文件路徑</param>
        /// <param name="picPath">浮水印圖片路徑</param>
        /// <returns>文件數據</returns>
        public async Task<byte[]> InsertWatermark2(string docPath, string picPath)
        {
            using FileStream fileStream = new(docPath, FileMode.Open);
            using MemoryStream memoryStream = new();
            await fileStream.CopyToAsync(memoryStream);
            memoryStream.Position = 0;

            using (WordprocessingDocument package = WordprocessingDocument.Open(memoryStream, true))
            {
                await InsertCustomWatermark2(package, picPath);
            }

            return memoryStream.ToArray();
        }

        /// <summary>
        /// 加入客製浮水印 - 直接讀取圖片byte[]版本
        /// </summary>
        /// <param name="package">文件主體</param>
        /// <param name="picPath">圖片路徑</param>
        /// <exception cref="Exception"></exception>
        private async Task InsertCustomWatermark2(WordprocessingDocument package, string picPath)
        {
            MainDocumentPart? mainDocumentPart = package.MainDocumentPart;

            if (mainDocumentPart != null)
            {
                // 刪除原本的HeaderParts
                mainDocumentPart.DeleteParts(mainDocumentPart.HeaderParts);

                // 創建新的HeaderParts
                HeaderPart headPart1 = mainDocumentPart.AddNewPart<HeaderPart>();

                // 在HeaderParts裡設定Header，繪製浮水印區塊
                GenerateHeaderPartContent(headPart1);

                // 取得該part的relationshipId
                string rId = mainDocumentPart.GetIdOfPart(headPart1);

                // 創建For Image的Part
                ImagePart image = headPart1.AddNewPart<ImagePart>("image/png", "rId999");

                // 輸入Image的數據
                await GenerateImagePartContent(image, picPath);

                SetHeaderPartReference(mainDocumentPart, rId);
            }
            else
            {
                throw new Exception("文件無內容");
            }
        }

        /// <summary>
        /// 填入圖片Part內容
        /// </summary>
        /// <param name="imagePart">ImagePart</param>
        /// <param name="picPath">圖片路徑</param>
        private async Task GenerateImagePartContent(ImagePart imagePart, string picPath)
        {
            using FileStream fileStream = new(picPath, FileMode.Open, FileAccess.Read);
            fileStream.Position = 0;
            using MemoryStream memoryStream = new();
            await fileStream.CopyToAsync(memoryStream);
            memoryStream.Position = 0;

            imagePart.FeedData(memoryStream);
        }
        #endregion

        /// <summary>
        /// 產生頁首內容
        /// </summary>
        /// <param name="headerPart">頁首Part</param>
        private void GenerateHeaderPartContent(HeaderPart headerPart)
        {
            Header header = new Header();
            Paragraph paragraph = new Paragraph();
            Run run = new Run();
            Picture picture = new Picture();
            V.Shape shape = new V.Shape()
            {
                Id = "WordPictureWatermark75517470",
                Style = "position:absolute;" +
                    "left:0;" +
                    "text-align:left;" +
                    "margin-left:300pt;" +
                    "margin-top:450pt;" +
                    "width:160pt;" +
                    "height:160pt;" +
                    "z-index:-251656192;",
                //"mso-position-horizontal:center;" +
                //"mso-position-horizontal-relative:margin;" +
                //"mso-position-vertical:center;" +
                //"mso-position-vertical-relative:margin",
                OptionalString = "_x0000_s2051",
                AllowInCell = false,
                Type = "#_x0000_t75"
            };
            V.ImageData imageData = new V.ImageData()
            {
                Gain = "1da1ce", // 鮮紅: 21dd7e
                BlackLevel = "200000",
                Title = "水印",
                RelationshipId = "rId999" // 目標ImagePartRId
            };
            Stroke stroke = new Stroke() { On = false }; // 關閉外框

            shape.AppendChild(stroke);
            shape.Append(imageData);
            picture.Append(shape);
            run.Append(picture);
            paragraph.Append(run);
            header.Append(paragraph);
            headerPart.Header = header;
        }

        /// <summary>
        /// 生成空的HeaderPart
        /// </summary>
        /// <param name="mainDocumentPart">MainDocumentPart</param>
        /// <returns>空的HeaderPart</returns>
        private HeaderPart CreateBlankHeader(MainDocumentPart mainDocumentPart)
        {
            var headerPart = mainDocumentPart.AddNewPart<HeaderPart>();
            Header header = new Header();
            Paragraph paragraph = new Paragraph();
            Run run = new Run();
            paragraph.Append(run);
            header.Append(paragraph);
            headerPart.Header = header;

            return headerPart;
        }

        /// <summary>
        /// 繫結MainDocumentPart與目標HeaderPart
        /// </summary>
        /// <param name="mainDocumentPart">MainDocumentPart</param>
        /// <param name="rId">目標HeaderPart rId</param>
        /// <remarks>最後一節之外皆加入水印</remarks>
        private void SetHeaderPartReference(MainDocumentPart mainDocumentPart, string rId)
        {
            // Body下的SectionProperties > HeaderReference 只控制最後一節
            // 其他的位於 Paragraph > SectionProperties > HeaderReference
            // 如果該節沒有訂義HeaderReference，會向上繼承
            // 所以最後尾如果不想要有浮水印，要自己創一個空的
            List<SectionProperties>? sectPrs = mainDocumentPart.Document.Body?.Descendants<SectionProperties>().ToList();

            for (int i = 0; i < sectPrs?.Count; i++)
            {
                var sectPr = sectPrs[i];

                sectPr.RemoveAllChildren<HeaderReference>();

                if (i != sectPrs.Count - 1)
                {
                    sectPr.PrependChild<HeaderReference>(new HeaderReference() { Type = HeaderFooterValues.First, Id = rId }); // 該節第一頁
                    sectPr.PrependChild<HeaderReference>(new HeaderReference() { Type = HeaderFooterValues.Default, Id = rId }); // 該節預設
                    sectPr.PrependChild<HeaderReference>(new HeaderReference() { Type = HeaderFooterValues.Even, Id = rId }); // 該節偶數
                }
                else
                {
                    var blankHeaderPart = CreateBlankHeader(mainDocumentPart);
                    string blankHeaderPartRid = mainDocumentPart.GetIdOfPart(blankHeaderPart);

                    sectPr.PrependChild<HeaderReference>(new HeaderReference() { Type = HeaderFooterValues.First, Id = blankHeaderPartRid });
                    sectPr.PrependChild<HeaderReference>(new HeaderReference() { Type = HeaderFooterValues.Default, Id = blankHeaderPartRid });
                    sectPr.PrependChild<HeaderReference>(new HeaderReference() { Type = HeaderFooterValues.Even, Id = blankHeaderPartRid });
                }
            }
        }
    }
}
