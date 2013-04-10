//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;

namespace AddingSlidesToPresentation
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            /***********************  Adding Empty Slide in a Presentation ******************************/

            //Instantiate a Presentation object that represents a PPT file
            Presentation pres = new Presentation(dataDir + "demo.ppt");
            
            //Adding an empty slide to the presentation and getting the reference of
            //that empty slide
            Slide slide1 = pres.AddEmptySlide();

            //Writing the presentation as a PPT file
            pres.Write(dataDir + "EmptySlide.ppt");

            /***********************  Adding Body Slide in a Presentation ******************************/

            //Instantiate a Presentation object that represents a PPT file
            pres = new Presentation(dataDir + "demo.ppt");

            //Adding a body slide to the presentation and getting the reference of
            //that body slide
            Slide slide2 = pres.AddBodySlide();

            //Writing the presentation as a PPT file
            pres.Write(dataDir + "BodySlide.ppt");

            /***********************  Adding Double Body Slide in a Presentation ******************************/

            //Instantiate a Presentation object that represents a PPT file
            pres = new Presentation(dataDir + "demo.ppt");
                        
            //Adding a double body slide to the presentation and getting the reference of
            //that double body slide
            Slide slide3 = pres.AddDoubleBodySlide();

            //Writing the presentation as a PPT file
            pres.Write(dataDir + "DoubleBodySlide.ppt");

            /***********************  Adding Header Slide in a Presentation ******************************/

            //Instantiate a Presentation object that represents a PPT file
            pres = new Presentation(dataDir + "demo.ppt");

            //Adding a header slide to the presentation and getting the reference of
            //that header slide
            Slide slide4 = pres.AddHeaderSlide();

            //Writing the presentation as a PPT file
            pres.Write(dataDir + "HeaderSlide.ppt");

            /***********************  Adding Title Slide in a Presentation ******************************/

            //Instantiate a Presentation object that represents a PPT file
            pres = new Presentation(dataDir + "demo.ppt");

            //Adding a title slide to the presentation and getting the reference of
            //that title slide
            Slide slide5 = pres.AddTitleSlide();
            
            //Writing the presentation as a PPT file
            pres.Write(dataDir + "TitleSlide.ppt");

            
        }
    }
}