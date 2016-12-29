using System.IO;
using Aspose.Slides;
using System;

namespace Aspose.Slides.Examples.CSharp.SmartArts
{
    public class AssistantNode
    {
        public static void Run()
        {
            // ExStart:AssistantNode
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_SmartArts();

            // Creating a presentation instance
            using (Presentation pres = new Presentation(dataDir+ "AssistantNode.pptx"))
            {
                // Traverse through every shape inside first slide
                foreach (IShape shape in pres.Slides[0].Shapes)
                {
                    // Check if shape is of SmartArt type
                    if (shape is Aspose.Slides.SmartArt.ISmartArt)
                    {
                        // Typecast shape to SmartArtEx
                        Aspose.Slides.SmartArt.ISmartArt smart = (Aspose.Slides.SmartArt.SmartArt)shape;
                        // Traversing through all nodes of SmartArt shape

                        foreach (Aspose.Slides.SmartArt.ISmartArtNode node in smart.AllNodes)
                        {
                            String tc = node.TextFrame.Text;
                            // Check if node is Assitant node
                            if (node.IsAssistant)
                            {
                                // Setting Assitant node to false and making it normal node
                                node.IsAssistant = false;
                            }
                        }
                    }
                }
                // Save Presentation
                pres.Save(dataDir + "ChangeAssitantNode_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            }
            // ExEnd:AssistantNode
        }
    }
}