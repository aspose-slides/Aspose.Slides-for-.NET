using System.Drawing;
using Aspose.Slides.SmartArt;
using Aspose.Slides.Export;
using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.SmartArts
{
    class FillFormatSmartArtShapeNode
    {
        public static void Run()
        {
            // ExStart:FillFormatSmartArtShapeNode
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_SmartArts();

            using (Presentation presentation = new Presentation())
            {
                // Accessing the slide
                ISlide slide = presentation.Slides[0];

                // Adding SmartArt shape and nodes
                var chevron = slide.Shapes.AddSmartArt(10, 10, 800, 60, SmartArtLayoutType.ClosedChevronProcess);
                var node = chevron.AllNodes.AddNode();
                node.TextFrame.Text = "Some text";

                // Setting node fill color
                foreach (var item in node.Shapes)
                {
                    item.FillFormat.FillType = FillType.Solid;
                    item.FillFormat.SolidFillColor.Color = Color.Red;
                }

                // Saving Presentation
                presentation.Save(dataDir + "FillFormat_SmartArt_ShapeNode_out.pptx", SaveFormat.Pptx);
            }
            // ExEnd:FillFormatSmartArtShapeNode
        }
    }
}
