using System.Drawing;
using Aspose.Slides.Export;
using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Tables
{
    public class TextFormattingInsideTableColumn
    {
        public static void Run()
        {
            // ExStart:TextFormattingInsideTableColumn
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Tables();

            // Create an instance of Presentation class
            Presentation pres = new Presentation();
           
            ISlide slide = pres.Slides[0];

            ITable someTable = pres.Slides[0].Shapes[0] as ITable; // let's say that the first shape on the first slide is a table

            // setting first column cells' font height
            PortionFormat portionFormat = new PortionFormat();
            portionFormat.FontHeight = 25;
            someTable.Columns[0].SetTextFormat(portionFormat);

            // setting first column cells' text alignment and right margin in one call
            ParagraphFormat paragraphFormat = new ParagraphFormat();
            paragraphFormat.Alignment = TextAlignment.Right;
            paragraphFormat.MarginRight = 20;
            someTable.Columns[0].SetTextFormat(paragraphFormat);

            // setting second column cells' text vertical type
            TextFrameFormat textFrameFormat = new TextFrameFormat();
            textFrameFormat.TextVerticalType = TextVerticalType.Vertical;
            someTable.Columns[1].SetTextFormat(textFrameFormat);

            pres.Save(path + "result.pptx", Aspose.Slides.Export.SaveFormat.Pptx);

            // ExEnd:TextFormattingInsideTableColumn
         }
    }
}

