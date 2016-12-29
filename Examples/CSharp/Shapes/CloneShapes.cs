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
    class CloneShapes
    {
        public static void Run()
        {
            //ExStart:CloneShapes
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Shapes();

            // Instantiate Presentation class
            using (Presentation srcPres = new Presentation(dataDir + "Source Frame.pptx"))
            {
                IShapeCollection sourceShapes = srcPres.Slides[0].Shapes;
                ILayoutSlide blankLayout = srcPres.Masters[0].LayoutSlides.GetByType(SlideLayoutType.Blank);
                ISlide destSlide = srcPres.Slides.AddEmptySlide(blankLayout);
                IShapeCollection destShapes = destSlide.Shapes;
                destShapes.AddClone(sourceShapes[1], 50, 150 + sourceShapes[0].Height);
                destShapes.AddClone(sourceShapes[2]);                 
                destShapes.InsertClone(0, sourceShapes[0], 50, 150);

                // Write the PPTX file to disk
                srcPres.Save(dataDir + "CloneShape_out.pptx", SaveFormat.Pptx);
                //ExStart:CloneShapes
            }
           
        }
    }
}




