 
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using AsposeSmartArts = Aspose.Slides.SmartArt;

namespace Aspose.Slides.Examples.CSharp.SmartArts

{
    public class AddNodes
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_SmartArts();

            //Load the desired the presentation//Load the desired the presentation
            Presentation pres = new Presentation(dataDir+ "AddNodes.pptx");

            //Traverse through every shape inside first slide
            foreach (IShape shape in pres.Slides[0].Shapes)
            {

                //Check if shape is of SmartArt type
                if (shape is AsposeSmartArts.SmartArt)
                {

                    //Typecast shape to SmartArt
                    AsposeSmartArts.SmartArt smart = (AsposeSmartArts.SmartArt)shape;

                    //Adding a new SmartArt Node
                    AsposeSmartArts.SmartArtNode TemNode = (AsposeSmartArts.SmartArtNode)smart.AllNodes.AddNode();

                    //Adding text
                    TemNode.TextFrame.Text = "Test";

                    //Adding new child node in parent node. It  will be added in the end of collection
                    AsposeSmartArts.SmartArtNode newNode = (AsposeSmartArts.SmartArtNode)TemNode.ChildNodes.AddNode();

                    //Adding text
                    newNode.TextFrame.Text = "New Node Added";

                }
            }

            //Saving Presentation
            pres.Save(dataDir+ "AddSmartArtNode.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            
            
        }
    }
}