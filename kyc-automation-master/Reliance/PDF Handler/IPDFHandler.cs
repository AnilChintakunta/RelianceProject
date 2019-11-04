using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reliance.PDF_Handler
{
    public interface IPDFHandler
    {
        List<string> GetPdfPagesContent(Stream pdfStream);

        void DeleteFile(string directoryPath);
    }
}
