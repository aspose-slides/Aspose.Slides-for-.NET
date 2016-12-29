using System.IO;
using Aspose.Slides;
using System;

namespace Aspose.Slides.Examples.CSharp.SmartArts
{
    public class AccessChildNodes
    {
        public static void Run()
        {
            // ExStart:AccessChildNodes
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_SmartArts();

            // Load the desired the presentation
            Presentation pres = new Presentation(dataDir+ "AccessChildNodes.pptx");

            // Traverse through every shape inside first slide
            foreach (IShape shape in pres.Slides[0].Shapes)
            {

                // Check if shape is of SmartArt type
                if (shape is Aspose.Slides.SmartArt.SmartArt)
                {

                    // Typecast shape to SmartArt
                    Aspose.Slides.SmartArt.SmartArt smart = (Aspose.Slides.SmartArt.SmartArt)shape;

                    // Traverse through all nodes inside SmartArt
                    for (int i = 0; i < smart.AllNodes.Count; i++)
                    {
                        // Accessing SmartArt node at index i
                        Aspose.Slides.SmartArt.SmartArtNode node0 = (Aspose.Slides.SmartArt.SmartArtNode)smart.AllNodes[i];

                        // Traversing through the child nodes in SmartArt node at index i
                        for (int j = 0; j < node0.ChildNodes.Count; j++)
                        {
                            // Accessing the child node in SmartArt node
                            Aspose.Slides.SmartArt.SmartArtNode node = (Aspose.Slides.SmartArt.SmartArtNode)node0.ChildNodes[j];

                            // Printing the SmartArt child node parameters
                            string outString = string.Format("j = {0}, Text = {1},  Level = {2}, Position = {3}", j, node.TextFrame.Text, node.Level, node.Position);
                            Console.WriteLine(outString);
                        }
                    }
                }
            }
            // ExEnd:AccessChildNodes
        }
    }
}