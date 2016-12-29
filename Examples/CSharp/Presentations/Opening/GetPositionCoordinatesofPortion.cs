using System;
using System.Drawing;
using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.Opening
{
    class GetPositionCoordinatesofPortion
    {
        public static void Run()
        {
            //ExStart:GetPositionCoordinatesofPortion
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_PresentationOpening();
            using (Presentation presentation = new Presentation(dataDir + "Shapes.pptx"))
            {
                IAutoShape shape = (IAutoShape)presentation.Slides[0].Shapes[0];
                var textFrame = (ITextFrame)shape.TextFrame;

                foreach (var paragraph in textFrame.Paragraphs)
                {
                    foreach (Portion portion in paragraph.Portions)
                    {
                        PointF point = portion.GetCoordinates();
                        Console.Write(Environment.NewLine + "Corrdinates X =" + point.X + " Corrdinates Y =" + point.Y);
                    }
                }
            }
            //ExEnd:GetPositionCoordinatesofPortion
        }
    }
}

