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
using System;

namespace CSharp.SmartArts
{
    public class AccessSmartArt
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_SmartArts();

            //Load the desired the presentation
            //Load the desired the presentation
            Presentation pres = new Presentation(dataDir + "AccessSmartArt.pptx");

            //Traverse through every shape inside first slide
            foreach (IShape shape in pres.Slides[0].Shapes)
            {

                //Check if shape is of SmartArt type
                if (shape is SmartArt)
                {

                    //Typecast shape to SmartArt
                    SmartArt smart = (SmartArt)shape;

                    //Traverse through all nodes inside SmartArt
                    for (int i = 0; i < smart.AllNodes.Count; i++)
                    {
                        //Accessing SmartArt node at index i
                        SmartArtNode node = (SmartArtNode)smart.AllNodes[i];

                        //Printing the SmartArt node parameters
                        string outString = string.Format("i = {0}, Text = {1},  Level = {2}, Position = {3}", i, node.TextFrame.Text, node.Level, node.Position);
                        Console.WriteLine(outString);
                    }
                }
            }
            
            
        }
    }
}