using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Reliance.PDF_Handler
{
    public class PDFHandler: IPDFHandler
    {
        private SimpleTextExtractionStrategy strategy;
        private IWebBrowser webBrowser;

        public PDFHandler(IWebBrowser webBrowser)
        {
            this.webBrowser = webBrowser;
        }

        public List<string> GetPdfPagesContent(Stream pdfStream)
        {
            List<string> allPDFPagesContent = new List<string>();
            PdfReader reader = new PdfReader(pdfStream);
            PdfReaderContentParser parser = new PdfReaderContentParser(reader);
            for (int i = 1; i <= reader.NumberOfPages; i++)
            {
                strategy = parser.ProcessContent(i, new SimpleTextExtractionStrategy());
                allPDFPagesContent.Add(strategy.GetResultantText());
            }

            return allPDFPagesContent;
        }

        public void DeleteFile(string directoryPath)
        {
            try
            {
                DirectoryInfo di = new DirectoryInfo(directoryPath);
                foreach (FileInfo file in di.GetFiles())
                {
                    file.Delete();
                }
            }
            catch (System.Exception)
            {
                Console.WriteLine("Make sure no files are opened, Please close all pdf files");
                throw;
            }
            
        }
    }
}
