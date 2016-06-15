using System.IO;
using Aspose.Slides;
using AsposeSlides = Aspose.Slides.SmartArt;
using System;

namespace Aspose.Slides.Examples.CSharp.SmartArts
{
    public class AssistantNode
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_SmartArts();

            //Creating a presentation instance
            using (Presentation pres = new Presentation(dataDir+ "AssistantNode.pptx"))
            {
                //Traverse through every shape inside first slide
                foreach (IShape shape in pres.Slides[0].Shapes)
                {
                    //Check if shape is of SmartArt type
                    if (shape is AsposeSlides.ISmartArt)
                    {
                        //Typecast shape to SmartArtEx
                        AsposeSlides.ISmartArt smart = (AsposeSlides.SmartArt)shape;
                        //Traversing through all nodes of SmartArt shape

                        foreach (AsposeSlides.ISmartArtNode node in smart.AllNodes)
                        {
                            String tc = node.TextFrame.Text;
                            //Check if node is Assitant node
                            if (node.IsAssistant)
                            {
                                //Setting Assitant node to false and making it normal node
                                node.IsAssistant = false;
                            }
                        }
                    }
                }
                //Save Presentation
                pres.Save(dataDir+ "ChangeAssitantNode.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            }            
        }
    }
}