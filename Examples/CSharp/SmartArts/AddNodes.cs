//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;
using Aspose.Slides.SmartArt;

namespace CSharp.SmartArts
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
                if (shape is SmartArt)
                {

                    //Typecast shape to SmartArt
                    SmartArt smart = (SmartArt)shape;

                    //Adding a new SmartArt Node
                    SmartArtNode TemNode = (SmartArtNode)smart.AllNodes.AddNode();

                    //Adding text
                    TemNode.TextFrame.Text = "Test";

                    //Adding new child node in parent node. It  will be added in the end of collection
                    SmartArtNode newNode = (SmartArtNode)TemNode.ChildNodes.AddNode();

                    //Adding text
                    newNode.TextFrame.Text = "New Node Added";

                }
            }

            //Saving Presentation
            pres.Save(dataDir+ "AddSmartArtNode.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            
            
        }
    }
}