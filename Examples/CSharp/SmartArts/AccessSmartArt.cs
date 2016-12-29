using System.IO;
using Aspose.Slides;
using System;

namespace Aspose.Slides.Examples.CSharp.SmartArts
{
    public class AccessSmartArt
    {
        public static void Run()
        {
            // ExStart:AccessSmartArt
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_SmartArts();

            // Load the desired the presentation
             Presentation pres = new Presentation(dataDir + "AccessSmartArt.pptx");

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
                        Aspose.Slides.SmartArt.SmartArtNode node = (Aspose.Slides.SmartArt.SmartArtNode)smart.AllNodes[i];

                        // Printing the SmartArt node parameters
                        string outString = string.Format("i = {0}, Text = {1},  Level = {2}, Position = {3}", i, node.TextFrame.Text, node.Level, node.Position);
                        Console.WriteLine(outString);
                    }
                }
            }
            // ExEnd:AccessSmartArt
        }
    }
}