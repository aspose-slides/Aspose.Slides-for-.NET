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
    public class RemoveNodeSpecificPosition
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_SmartArts();

            //Load the desired the presentation             
            Presentation pres = new Presentation(dataDir + "RemoveNodeSpecificPosition.pptx");

            //Traverse through every shape inside first slide
            foreach (IShape shape in pres.Slides[0].Shapes)
            {

                //Check if shape is of SmartArt type
                if (shape is SmartArt)
                {
                    //Typecast shape to SmartArt
                    SmartArt smart = (SmartArt)shape;

                    if (smart.AllNodes.Count > 0)
                    {
                        //Accessing SmartArt node at index 0
                        ISmartArtNode node = smart.AllNodes[0];

                        if (node.ChildNodes.Count >= 2)
                        {
                            //Removing the child node at position 1
                            ((SmartArtNodeCollection)node.ChildNodes).RemoveNodeByPosition(1);
                        }

                    }
                }

            }

            //Save Presentation
            pres.Save(dataDir+ "RemoveSmartArtNodeByPosition.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            
        }
    }
}