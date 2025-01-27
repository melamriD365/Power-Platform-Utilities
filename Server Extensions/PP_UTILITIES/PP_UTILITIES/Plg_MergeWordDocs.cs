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
                if (context.InputParameters.Contains("meaf_docs") && context.InputParameters["meaf_docs"] is string[])
                {
                    string[] docsBase64 = (string[])context.InputParameters["meaf_docs"];

                    if (docsBase64 == null || docsBase64.Length < 2)
                    {
                        throw new InvalidPluginExecutionException("Input documents cannot be null or fewer than two.");
                    }

                    bool addPageBreak = true;
                    if (context.InputParameters.Contains("meaf_AddPageBreak") && context.InputParameters["meaf_AddPageBreak"] is bool)
                    {  
                        addPageBreak = (bool)context.InputParameters["meaf_AddPageBreak"];
                    }

                    var docBytesList = docsBase64.Select(Convert.FromBase64String).ToArray();
                    byte[] mergedDocBytes = MergeWordDocs(docBytesList, addPageBreak);
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

        private byte[] MergeWordDocs(byte[][] docs, bool addPageBreak)
        {
            byte[] destBytes = docs[0];
            for (int i = 1; i < docs.Length; i++)
            {
                destBytes = Merge(destBytes, docs[i], addPageBreak);
            }
            return destBytes;
        }

        private byte[] Merge(byte[] dest, byte[] src, bool addPageBreak)
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

                        if (addPageBreak)
                        {
                            var pageBreak = new Paragraph(new Run(new Break { Type = BreakValues.Page }));
                            mainPart.Document.Body.AppendChild(pageBreak);
                        }

                        var altChunk = new AltChunk { Id = altChunkId };
                        mainPart.Document.Body.AppendChild(altChunk);

                        mainPart.Document.Save();
                    }
                }
                return memoryStreamDest.ToArray();
            }
        }
    }
}
