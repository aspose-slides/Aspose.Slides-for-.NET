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

namespace CSharp.SmartArts
{
    public class AccessSmartArtShape
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_SmartArts();

            //Load the desired the presentation
            using (Presentation pres = new Presentation(dataDir+ "AccessSmartArtShape.pptx"))
            {

                //Traverse through every shape inside first slide
                foreach (IShape shape in pres.Slides[0].Shapes)
                {
                    //Check if shape is of SmartArt type
                    if (shape is ISmartArt)
                    {
                        //Typecast shape to SmartArtEx
                        ISmartArt smart = (ISmartArt)shape;
                        System.Console.WriteLine("Shape Name:"+ smart.Name);
                    
                    }
                }

            }

            
            
        }
    }
}