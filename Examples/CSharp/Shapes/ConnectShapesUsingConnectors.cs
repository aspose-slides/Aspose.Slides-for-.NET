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
    class ConnectShapesUsingConnectors
    {
        public static void Run()
        {
            //ExStart:ConnectShapesUsingConnectors
            // The path to the documents directory.                    
            string dataDir = RunExamples.GetDataDir_Shapes();
            
            // Instantiate Presentation class that represents the PPTX file
            using (Presentation input = new Presentation())
            {                
                // Accessing shapes collection for selected slide
                IShapeCollection shapes = input.Slides[0].Shapes;

                // Add autoshape Ellipse
                IAutoShape ellipse = shapes.AddAutoShape(ShapeType.Ellipse, 0, 100, 100, 100);

                // Add autoshape Rectangle
                IAutoShape rectangle = shapes.AddAutoShape(ShapeType.Rectangle, 100, 300, 100, 100);

                // Adding connector shape to slide shape collection
                IConnector connector = shapes.AddConnector(ShapeType.BentConnector2, 0, 0, 10, 10);

                // Joining Shapes to connectors
                connector.StartShapeConnectedTo = ellipse;
                connector.EndShapeConnectedTo = rectangle;

                // Call reroute to set the automatic shortest path between shapes
                connector.Reroute();

                // Saving presenation
                input.Save(dataDir + "Connecting shapes using connectors_out.pptx", SaveFormat.Pptx);
            }
            //ExEnd:ConnectShapesUsingConnectors
        }
    }
}
