using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Presentations.Opening
{
    class SetAccessPermissionsToPDF
    {

        public static void Run() {

            //ExStart:SetAccessPermissionsToPDF

            string dataDir = RunExamples.GetDataDir_PresentationOpening();

            var pdfOptions = new PdfOptions();
            pdfOptions.Password = "my_password";
            pdfOptions.AccessPermissions = PdfAccessPermissions.PrintDocument | PdfAccessPermissions.HighQualityPrint;

            using (var presentation = new Presentation())
            {
                presentation.Save(dataDir + "PDFWithPermissions.pdf", SaveFormat.Pdf, pdfOptions);
            }
            //ExEnd:SetAccessPermissionsToPDF


        }
    }
}
