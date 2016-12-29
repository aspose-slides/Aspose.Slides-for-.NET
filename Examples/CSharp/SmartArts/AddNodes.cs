 using System.IO;
using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
namespace Aspose.Slides.Examples.CSharp.SmartArts

{
    public class AddNodes
    {
        public static void Run()
        {
            // ExStart:AddNodes
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_SmartArts();

            // Load the desired the presentation// Load the desired the presentation
            Presentation pres = new Presentation(dataDir+ "AddNodes.pptx");

            // Traverse through every shape inside first slide
            foreach (IShape shape in pres.Slides[0].Shapes)
            {

                // Check if shape is of SmartArt type
                if (shape is Aspose.Slides.SmartArt.SmartArt)
                {

                    // Typecast shape to SmartArt
                    Aspose.Slides.SmartArt.SmartArt smart = (Aspose.Slides.SmartArt.SmartArt)shape;

                    // Adding a new SmartArt Node
                    Aspose.Slides.SmartArt.SmartArtNode TemNode = (Aspose.Slides.SmartArt.SmartArtNode)smart.AllNodes.AddNode();

                    // Adding text
                    TemNode.TextFrame.Text = "Test";

                    // Adding new child node in parent node. It  will be added in the end of collection
                    Aspose.Slides.SmartArt.SmartArtNode newNode = (Aspose.Slides.SmartArt.SmartArtNode)TemNode.ChildNodes.AddNode();

                    // Adding text
                    newNode.TextFrame.Text = "New Node Added";

                }
            }

            // Saving Presentation
            pres.Save(dataDir + "AddSmartArtNode_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            // ExEnd:AddNodes
        }
    }
}