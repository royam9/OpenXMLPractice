using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using V = DocumentFormat.OpenXml.Vml;
using System.Diagnostics;

namespace Services
{
    public class WordWatermarkService
    {
        public byte[] InsertWatermark(string docPath, string picPath)
        {
            using FileStream fileStream = new(docPath, FileMode.Open);
            using MemoryStream memoryStream = new();
            fileStream.CopyTo(memoryStream);
            memoryStream.Position = 0;

            using (WordprocessingDocument package = WordprocessingDocument.Open(memoryStream, true))
            {
                InsertCustomWatermark(package, picPath);
            }

            return memoryStream.ToArray();
        }

        private void InsertCustomWatermark(WordprocessingDocument package, string p)
        {
            // 取得圖片的base64字串
            SetWaterMarkPicture(p);
            MainDocumentPart mainDocumentPart1 = package.MainDocumentPart;
            if (mainDocumentPart1 != null)
            {
                // 刪除原本的HeaderParts
                mainDocumentPart1.DeleteParts(mainDocumentPart1.HeaderParts);
                // 新增新的HeaderParts
                HeaderPart headPart1 = mainDocumentPart1.AddNewPart<HeaderPart>();
                // 在HeaderParts裡設定Header，繪製浮水印區塊
                GenerateHeaderPart1Content(headPart1);
                // 取得該part的relationshipId
                string rId = mainDocumentPart1.GetIdOfPart(headPart1);
                // 新增For Image的Part
                ImagePart image = headPart1.AddNewPart<ImagePart>("image/png", "rId999");
                // 輸入Image的數據
                GenerateImagePart1Content(image);
                // 透過段落的HeaderReference指向目標HeaderPart
                //IEnumerable<SectionProperties> sectPrs = mainDocumentPart1.Document.Body.Elements<SectionProperties>();
                //foreach (var sectPr in sectPrs)
                //{
                //    // sectPr.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

                //    sectPr.RemoveAllChildren<HeaderReference>();
                //    sectPr.PrependChild<HeaderReference>(new HeaderReference() { Type = HeaderFooterValues.Even, Id = rId });
                //    sectPr.PrependChild<HeaderReference>(new HeaderReference() { Type = HeaderFooterValues.Default, Id = rId });
                //    sectPr.PrependChild<HeaderReference>(new HeaderReference() { Type = HeaderFooterValues.First, Id = rId });
                //}

                IEnumerable<SectionProperties> sectPrs = mainDocumentPart1.Document.Body.Descendants<SectionProperties>();
                foreach (var sectPr in sectPrs)
                {
                    sectPr.RemoveAllChildren<HeaderReference>();
                    sectPr.PrependChild<HeaderReference>(new HeaderReference() { Type = HeaderFooterValues.Even, Id = rId });
                    sectPr.PrependChild<HeaderReference>(new HeaderReference() { Type = HeaderFooterValues.Default, Id = rId });
                    sectPr.PrependChild<HeaderReference>(new HeaderReference() { Type = HeaderFooterValues.First, Id = rId });
                }
            }
            //else
            //{
            //    MessageBox.Show("alert");
            //}
        }

        private void GenerateHeaderPart1Content(HeaderPart headerPart1)
        {
            Header header1 = new Header();
            Paragraph paragraph2 = new Paragraph();
            Run run1 = new Run();
            Picture picture1 = new Picture();
            V.Shape shape1 = new V.Shape()
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
            V.ImageData imageData1 = new V.ImageData()
            {
                Gain = "1da1ce", // 鮮紅: 21dd7e
                BlackLevel = "200000",
                Title = "水印",
                RelationshipId = "rId999"
            };
            shape1.Append(imageData1);
            picture1.Append(shape1);
            run1.Append(picture1);
            paragraph2.Append(run1);
            header1.Append(paragraph2);
            headerPart1.Header = header1;
        }

        private void GenerateImagePart1Content(ImagePart imagePart1)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart1Data);
            imagePart1.FeedData(data);
            data.Close();
        }

        private string imagePart1Data = "";

        private System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }

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
                imagePart1Data = Convert.ToBase64String(byteArray, 0, byteArray.Length);
            }
            catch (Exception ex)
            {
                Debug.Print(ex.Message);
            }
        }
    }
}
