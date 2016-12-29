using System.IO;
using Aspose.Slides;
using Aspose.Slides.SmartArt;
using System;

namespace Aspose.Slides.Examples.CSharp.SmartArts
{
    public class AccessChildNodeSpecificPosition
    {
        public static void Run()
        {
            // ExStart:AccessChildNodeSpecificPosition
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_SmartArts();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            // Instantiate the presentation
            Presentation pres = new Presentation();

            // Accessing the first slide
            ISlide slide = pres.Slides[0];

            // Adding the SmartArt shape in first slide
            ISmartArt smart = slide.Shapes.AddSmartArt(0, 0, 400, 400, SmartArtLayoutType.StackedList);

            // Accessing the SmartArt  node at index 0
            ISmartArtNode node = smart.AllNodes[0];

            // Accessing the child node at position 1 in parent node
            int position = 1;
            SmartArtNode chNode = (SmartArtNode)node.ChildNodes[position]; 

            // Printing the SmartArt child node parameters
            string outString = string.Format("j = {0}, Text = {1},  Level = {2}, Position = {3}", position, chNode.TextFrame.Text, chNode.Level, chNode.Position);
            Console.WriteLine(outString);
            // ExEnd:AccessChildNodeSpecificPosition
        }
    }
}