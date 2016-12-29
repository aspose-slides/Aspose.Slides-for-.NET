using System.IO;
using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;

namespace Aspose.Slides.Examples.CSharp.SmartArts
{
    public class RemoveNodeSpecificPosition
    {
        public static void Run()
        {
            // ExStart:RemoveNodeSpecificPosition
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_SmartArts();

            // Load the desired the presentation             
            Presentation pres = new Presentation(dataDir + "RemoveNodeSpecificPosition.pptx");

            // Traverse through every shape inside first slide
            foreach (IShape shape in pres.Slides[0].Shapes)
            {
                // Check if shape is of SmartArt type
                if (shape is Aspose.Slides.SmartArt.SmartArt)
                {
                    // Typecast shape to SmartArt
                    Aspose.Slides.SmartArt.SmartArt smart = (Aspose.Slides.SmartArt.SmartArt)shape;

                    if (smart.AllNodes.Count > 0)
                    {
                        // Accessing SmartArt node at index 0
                        Aspose.Slides.SmartArt.ISmartArtNode node = smart.AllNodes[0];

                        if (node.ChildNodes.Count >= 2)
                        {
                            // Removing the child node at position 1
                            ((Aspose.Slides.SmartArt.SmartArtNodeCollection)node.ChildNodes).RemoveNode(1);
                        }

                    }
                }
            }

            // Save Presentation
            pres.Save(dataDir + "RemoveSmartArtNodeByPosition_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            // ExEnd:RemoveNodeSpecificPosition
        }
    }
}