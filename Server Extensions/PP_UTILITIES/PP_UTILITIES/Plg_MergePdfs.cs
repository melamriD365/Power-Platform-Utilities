using Microsoft.Xrm.Sdk;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.IO;

namespace PP_UTILITIES
{
    public class Plg_MergePdfs : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            try
            {
                // Retrieve the input parameter
                if (context.InputParameters.Contains("meaf_pdfs") && context.InputParameters["meaf_pdfs"] is string[])
                {
                    // Retrieve the array of base64-encoded PDF strings
                    string[] pdfsBase64 = (string[])context.InputParameters["meaf_pdfs"];

                    if (pdfsBase64 == null || pdfsBase64.Length == 0)
                    {
                        throw new InvalidPluginExecutionException("Input PDFs cannot be null or empty.");
                    }

                    // Convert base64 strings to byte arrays
                    var pdfBytesList = new byte[pdfsBase64.Length][];
                    for (int i = 0; i < pdfsBase64.Length; i++)
                    {
                        if (string.IsNullOrEmpty(pdfsBase64[i]))
                        {
                            throw new InvalidPluginExecutionException($"PDF at index {i} is null or empty.");
                        }
                        pdfBytesList[i] = Convert.FromBase64String(pdfsBase64[i]);
                    }

                    // Merge PDFs
                    byte[] mergedPdfBytes = MergePdfs(pdfBytesList);

                    // Encode merged PDF as Base64 string and set output
                    string mergedPdfBase64 = Convert.ToBase64String(mergedPdfBytes);
                    context.OutputParameters["meaf_mergedpdf"] = mergedPdfBase64;
                }
                else
                {
                    throw new InvalidPluginExecutionException("Required input parameter 'meaf_pdfs' is missing or invalid.");
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace($"Error in Merge PDF Plugin: {ex.Message}");
                throw new InvalidPluginExecutionException($"Error in Merge PDF Plugin: {ex.Message}", ex);
            }
        }

        private byte[] MergePdfs(byte[][] pdfs)
        {
            using (var outputDocument = new PdfDocument())
            {
                foreach (var pdfBytes in pdfs)
                {
                    using (var pdfStream = new MemoryStream(pdfBytes))
                    {
                        var inputDocument = PdfReader.Open(pdfStream, PdfDocumentOpenMode.Import);
                        CopyPages(inputDocument, outputDocument);
                    }
                }

                using (var outputStream = new MemoryStream())
                {
                    outputDocument.Save(outputStream);
                    return outputStream.ToArray();
                }
            }
        }

        private void CopyPages(PdfDocument from, PdfDocument to)
        {
            for (int i = 0; i < from.PageCount; i++)
            {
                to.AddPage(from.Pages[i]);
            }
        }
    }
}
