using System;
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
    class RemoveShape
    {
        public static void Run()
        {
            //ExStart:RemoveShape
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Shapes();

            // Create Presentation object
            Presentation pres = new Presentation();

            // Get the first slide
            ISlide sld = pres.Slides[0];

            // Add autoshape of rectangle type
            IShape shp1 = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 40, 150, 50);
            IShape shp2 = sld.Shapes.AddAutoShape(ShapeType.Moon, 160, 40, 150, 50);
            String alttext = "User Defined";
            int iCount = sld.Shapes.Count;
            for (int i = 0; i < iCount; i++)
            {
                AutoShape ashp = (AutoShape)sld.Shapes[0];
                if (String.Compare(ashp.AlternativeText, alttext, StringComparison.Ordinal) == 0)
                {
                    sld.Shapes.Remove(ashp);
                }
            }

            // Save presentation to disk
            pres.Save(dataDir + "RemoveShape_out.pptx", SaveFormat.Pptx);
            //ExEnd:RemoveShape
        }
    }
}


