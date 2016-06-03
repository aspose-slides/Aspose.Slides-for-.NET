using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace CSharp.ProgrammersGuide.Shapes
{
    class ConnectShapesUsingConnectors
    {
        public static void Run()
        {
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
                input.Save(dataDir + "Connecting shapes using connectors.pptx", SaveFormat.Pptx);
            }
        }
    }
}
