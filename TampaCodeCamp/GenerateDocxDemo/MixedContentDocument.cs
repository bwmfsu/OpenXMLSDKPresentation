using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using System.Text;

namespace GenerateDocxDemo
{

    /// <summary>
    /// This class is responsible for building up a document that includes html and WordprocessingML content.
    /// </summary>
    public class MixedContentDocument
    {
        private readonly DocumentSpecification _spec;

        public MixedContentDocument(DocumentSpecification spec)
        {
            _spec = spec;
        }

        /// <summary>
        /// Builds the document.
        /// </summary>
        /// <returns>A byte array containing the content of the document. </returns>
        public byte[] GenerateDocument()
        {
            byte[] docBytes;
            using (MemoryStream memStream = new MemoryStream())
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = doc.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    mainPart.Document.Body = new Body();

                    string html = string.Format("<html><body>{0}</body></html>", _spec.HTMLContent);

                    AlternativeFormatImportPart altPart = null;
                    AlternativeFormatImportPart altPart1 = null;

                    using (MemoryStream htmlStream = new MemoryStream())
                    {
                        byte[] data = Encoding.UTF8.GetBytes(html);
                        htmlStream.Write(data, 0, data.Length);
                        htmlStream.Position = 0;
                        altPart = mainPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html);
                        altPart.FeedData(htmlStream);
                    }

                    using (MemoryStream uploadedDocStream = new MemoryStream())
                    {
                        byte[] data = _spec.DocumentToMerge;
                        uploadedDocStream.Write(data, 0, data.Length);
                        uploadedDocStream.Position = 0;
                        altPart1 = mainPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.WordprocessingML);
                        altPart1.FeedData(uploadedDocStream);
                    }

                    string altPartId = mainPart.GetIdOfPart(altPart);
                    string altPartId1 = mainPart.GetIdOfPart(altPart1);

                    mainPart.Document.Body.Append(new AltChunk() { Id = altPartId });
                    mainPart.Document.Body.Append(new AltChunk() { Id = altPartId1 });
                    mainPart.Document.Save();
                }
                docBytes = memStream.ToArray();
            }

            return docBytes;
        }
    }
}
