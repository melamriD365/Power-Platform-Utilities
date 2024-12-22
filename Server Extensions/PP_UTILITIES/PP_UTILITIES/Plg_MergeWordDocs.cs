using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Xrm.Sdk;

namespace PP_UTILITIES
{
    public class Plg_MergeWordDocs : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            try
            {
                // Retrieve the input parameter
                if (context.InputParameters.Contains("meaf_docs") && context.InputParameters["meaf_docs"] is string[])
                {
                    // Retrieve the array of base64-encoded Word document strings
                    string[] docsBase64 = (string[])context.InputParameters["meaf_docs"];

                    if (docsBase64 == null || docsBase64.Length < 2)
                    {
                        throw new InvalidPluginExecutionException("Input documents cannot be null or fewer than two.");
                    }

                    // Convert base64 strings to byte arrays
                    var docBytesList = docsBase64.Select(Convert.FromBase64String).ToArray();

                    // Merge Word documents
                    byte[] mergedDocBytes = MergeWordDocs(docBytesList);

                    // Encode merged document as Base64 string and set output
                    string mergedDocBase64 = Convert.ToBase64String(mergedDocBytes);
                    context.OutputParameters["meaf_MergedDoc"] = mergedDocBase64;
                }
                else
                {
                    throw new InvalidPluginExecutionException("Required input parameter 'meaf_docs' is missing or invalid.");
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace($"Error in Merge Word Docs Plugin: {ex.Message}");
                throw new InvalidPluginExecutionException($"Error in Merge Word Docs Plugin: {ex.Message}", ex);
            }
        }

        private byte[] MergeWordDocs(byte[][] docs)
        {
            // Initialize the destination document with the first document
            byte[] destBytes = docs[0];
            for (int i = 1; i < docs.Length; i++)
            {
                destBytes = Merge(destBytes, docs[i]);
            }
            return destBytes;
        }

        private byte[] Merge(byte[] dest, byte[] src)
        {
            string altChunkId = "AltChunkId" + DateTime.Now.Ticks.ToString();

            using (var memoryStreamDest = new MemoryStream())
            {
                memoryStreamDest.Write(dest, 0, dest.Length);
                memoryStreamDest.Seek(0, SeekOrigin.Begin);
                using (var memoryStreamSrc = new MemoryStream(src))
                {
                    using (WordprocessingDocument destDoc = WordprocessingDocument.Open(memoryStreamDest, true))
                    {
                        var mainPart = destDoc.MainDocumentPart;
                        var altPart = mainPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.WordprocessingML, altChunkId);
                        altPart.FeedData(memoryStreamSrc);

                        var altChunk = new AltChunk { Id = altChunkId };

                        // Add a page break before inserting the new document
                        var pageBreak = new Paragraph(new Run(new Break { Type = BreakValues.Page }));
                        mainPart.Document.Body.AppendChild(pageBreak);
                        mainPart.Document.Body.AppendChild(altChunk);

                        // Save changes to the destination document
                        mainPart.Document.Save();
                    }
                }
                return memoryStreamDest.ToArray();
            }
        }
    }
}
