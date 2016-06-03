using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace CSharp.ProgrammersGuide.Shapes
{
    class ConnectShapeUsingConnectionSite
    {
        public static void Run()
        {
            // The path to the documents directory.                    
            string dataDir = RunExamples.GetDataDir_Shapes();

            // Instantiate Presentation class that represents the PPTX file
            using (Presentation presentation = new Presentation())
            {
                // Accessing shapes collection for selected slide
                IShapeCollection shapes = presentation.Slides[0].Shapes;

                // Adding connector shape to slide shape collection
                IConnector connector = shapes.AddConnector(ShapeType.BentConnector3, 0, 0, 10, 10);

                // Add autoshape Ellipse
                IAutoShape ellipse = shapes.AddAutoShape(ShapeType.Ellipse, 0, 100, 100, 100);

                // Add autoshape Rectangle
                IAutoShape rectangle = shapes.AddAutoShape(ShapeType.Rectangle, 100, 200, 100, 100);

                // Joining Shapes to connectors
                connector.StartShapeConnectedTo = ellipse;
                connector.EndShapeConnectedTo = rectangle;

                // Setting the desired connection site index of Ellipse shape for connector to get connected
                uint wantedIndex = 6;

                // Checking if desired index is less than maximum site index count
                if (ellipse.ConnectionSiteCount > wantedIndex)
                {
                    // Setting the desired connection site for connector on Ellipse
                    connector.StartShapeConnectionSiteIndex = wantedIndex;
                }

                // Save presentation
                presentation.Save(dataDir + "Connecting_Shape_on_desired_connection_site.pptx", SaveFormat.Pptx);
            }

        }
    }
}
