//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;
using Aspose.Slides.SmartArt;
using System;

namespace AssistantNode
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Creating a presentation instance
            using (Presentation pres = new Presentation(dataDir+ "SimpleSmartArt.pptx"))
            {
                //Traverse through every shape inside first slide
                foreach (IShape shape in pres.Slides[0].Shapes)
                {
                    //Check if shape is of SmartArt type
                    if (shape is ISmartArt)
                    {
                        //Typecast shape to SmartArtEx
                        ISmartArt smart = (SmartArt)shape;
                        //Traversing through all nodes of SmartArt shape

                        foreach (ISmartArtNode node in smart.AllNodes)
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
            //Creating a presentation instance
            
            
        }
    }
}