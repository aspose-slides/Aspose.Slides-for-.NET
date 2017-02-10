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
    public class SetTextFormattingInsideTable
    {
        public static void Run()
        {
            // ExStart:TextFormattingInsideTable
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Tables();

            // Create an instance of Presentation class
            Presentation presentation = new Presentation();
      
            ISlide slide = presentation.Slides[0];

            ITable someTable = presentation.Slides[0].Shapes[0] as ITable; // let's say that the first shape on the first slide is a table

            // setting table cells' font height
            PortionFormat portionFormat = new PortionFormat();
            portionFormat.FontHeight = 25;
            someTable.SetTextFormat(portionFormat);

            // setting table cells' text alignment and right margin in one call
            ParagraphFormat paragraphFormat = new ParagraphFormat();
            paragraphFormat.Alignment = TextAlignment.Right;
            paragraphFormat.MarginRight = 20;
            someTable.SetTextFormat(paragraphFormat);

            // setting table cells' text vertical type
            TextFrameFormat textFrameFormat = new TextFrameFormat();
            textFrameFormat.TextVerticalType = TextVerticalType.Vertical;
            someTable.SetTextFormat(textFrameFormat);


            presentation.Save(path + "result.pptx", Aspose.Slides.Export.SaveFormat.Pptx);


            // ExEnd:SetTextFormattingInsideTable
         }
    }
}

