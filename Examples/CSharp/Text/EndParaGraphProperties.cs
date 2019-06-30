using System.Drawing;
using Aspose.Slides;
using Aspose.Slides.Export;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Text
{
    class EndParaGraphProperties
    {
        public static void Run()
        {
            //ExStart:EndParaGraphProperties
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();
            using (Presentation pres = new Presentation(dataDir+"presentation.pptx"))
        {
             IAutoShape shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 200, 250);

             Paragraph para1 = new Paragraph();
             para1.Portions.Add(new Portion("Sample text"));

             Paragraph para2 = new Paragraph();
             para2.Portions.Add(new Portion("Sample text 2"));
             PortionFormat endParagraphPortionFormat = new PortionFormat();
             endParagraphPortionFormat.FontHeight = 48;
             endParagraphPortionFormat.LatinFont = new FontData("Times New Roman");
             para2.EndParagraphPortionFormat = endParagraphPortionFormat;

             shape.TextFrame.Paragraphs.Add(para1);
             shape.TextFrame.Paragraphs.Add(para2);

            pres.Save(dataDir+"pres.pptx", SaveFormat.Pptx);
            }
            }
            //ExEnd:EndParaGraphProperties
        }
    }

