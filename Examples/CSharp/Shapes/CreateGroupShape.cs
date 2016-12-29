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
    class CreateGroupShape
    {
        public static void Run()
        {
            //ExStart:CreateScalingFactorThumbnail
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Shapes();

            // Instantiate Prseetation class 
            using (Presentation pres = new Presentation())
            {
                // Get the first slide 
                ISlide sld = pres.Slides[0];

                // Accessing the shape collection of slides 
                IShapeCollection slideShapes = sld.Shapes;

                // Adding a group shape to the slide 
                IGroupShape groupShape = slideShapes.AddGroupShape();

                // Adding shapes inside added group shape 
                groupShape.Shapes.AddAutoShape(ShapeType.Rectangle, 300, 100, 100, 100);
                groupShape.Shapes.AddAutoShape(ShapeType.Rectangle, 500, 100, 100, 100);
                groupShape.Shapes.AddAutoShape(ShapeType.Rectangle, 300, 300, 100, 100);
                groupShape.Shapes.AddAutoShape(ShapeType.Rectangle, 500, 300, 100, 100);

                // Adding group shape frame 
                groupShape.Frame = new ShapeFrame(100, 300, 500, 40, NullableBool.False, NullableBool.False, 0);

                // Write the PPTX file to disk 
                pres.Save(dataDir + "GroupShape_out.pptx", SaveFormat.Pptx);
            }
            //ExEnd:CreateScalingFactorThumbnail
        }
    }
}



