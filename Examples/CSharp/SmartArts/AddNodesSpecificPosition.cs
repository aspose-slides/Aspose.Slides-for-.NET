
using System.IO;

using Aspose.Slides;
using Aspose.Slides.SmartArt;

namespace Aspose.Slides.Examples.CSharp.SmartArts
{
    public class AddNodesSpecificPosition
    {
        public static void Run()
        {
            // ExStart:AddNodesSpecificPosition
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_SmartArts();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            // Creating a presentation instance
            Presentation pres = new Presentation();

            // Access the presentation slide
            ISlide slide = pres.Slides[0];

            // Add Smart Art IShape
            ISmartArt smart = slide.Shapes.AddSmartArt(0, 0, 400, 400, SmartArtLayoutType.StackedList);

            // Accessing the SmartArt node at index 0
            ISmartArtNode node = smart.AllNodes[0];

            // Adding new child node at position 2 in parent node
            SmartArtNode chNode = (SmartArtNode)((SmartArtNodeCollection)node.ChildNodes).AddNodeByPosition(2);

            // Add Text
            chNode.TextFrame.Text = "Sample Text Added";

            // Save Presentation
            pres.Save(dataDir + "AddSmartArtNodeByPosition_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            // ExEnd:AddNodesSpecificPosition
        }
    }
}