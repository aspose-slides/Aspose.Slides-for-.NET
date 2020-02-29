using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Aspose.Slides;
using Aspose.Slides.Export;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
Install it and then add its reference to this project. For any issues, questions or suggestions 
Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Text
{
    public class SaveWithDefaaultRegularFont
    {
        /// <summary>
        /// The code below demonstrates saving presentation to Html and Pdf with different default regular font.
        /// </summary>
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Text();
            string outPath = RunExamples.OutPath;

            using (Presentation pres = new Presentation(dataDir + "SaveOptionsDefaultRegularFont.pptx"))
            {
                HtmlOptions htmlOpts = new HtmlOptions();
                htmlOpts.DefaultRegularFont = "Arial Black";
                pres.Save(outPath + "Presentation-out-ArialBlack.html", SaveFormat.Html, htmlOpts);
                htmlOpts.DefaultRegularFont = "Lucida Console";
                pres.Save(outPath + "Presentation-out-LucidaConsole.html", SaveFormat.Html, htmlOpts);

                PdfOptions pdfOpts = new PdfOptions();
                pdfOpts.DefaultRegularFont = "Arial Black";
                pres.Save(outPath + "Presentation-out-ArialBlack.pdf", SaveFormat.Pdf, pdfOpts);
            }
        }
    }
}