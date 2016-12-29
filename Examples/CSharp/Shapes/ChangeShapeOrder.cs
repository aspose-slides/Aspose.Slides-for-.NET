using Aspose.Slides.Export;
using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    class ChangeShapeOrder
    {
        public static void Run()
        {
            //ExStart:ChangeShapeOrder
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Shapes();

            Presentation presentation1 = new Presentation(dataDir + "HelloWorld.pptx");
            ISlide slide = presentation1.Slides[0];
            IAutoShape shp3 = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 200, 365, 400, 150);
            shp3.FillFormat.FillType = FillType.NoFill;
            shp3.AddTextFrame(" ");

            ITextFrame txtFrame = shp3.TextFrame;
            IParagraph para = txtFrame.Paragraphs[0];
            IPortion portion = para.Portions[0];
            portion.Text="Watermark Text Watermark Text Watermark Text";
            shp3 = slide.Shapes.AddAutoShape(ShapeType.Triangle, 200, 365, 400, 150);
            slide.Shapes.Reorder(2, shp3);
            presentation1.Save(dataDir + "Reshape_out.pptx", SaveFormat.Pptx);
            //ExEnd:ChangeShapeOrder
        }
    }
}


