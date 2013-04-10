//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;

namespace CloningASlide
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            /********************************** Clone a Slide in Same Presentation ************************************/
            
            //Instantiate a Presentation object that represents a PPT file
            Presentation pres1 = new Presentation(dataDir + "demo.ppt");


            //Accessing a slide using its slide position
            Slide slide = pres1.GetSlideByPosition(1);


            //Cloning the selected slide at the end of the same presentation file
            pres1.CloneSlide(slide, pres1.Slides.LastSlidePosition + 1);


            //Writing the presentation as a PPT file
            pres1.Write(dataDir + "CloneSlide1.ppt");

            //Instantiate a Presentation where the cloned slide will be added
            Presentation pres2 = new Presentation(dataDir + "demo2.ppt");


            //Creating SortedList object that is used to store the temporary information
            //about the masters of PPT file. No value should be added to it.
            System.Collections.SortedList sList = new System.Collections.SortedList();


            //Cloning the selected slide at the end of another presentation file
            pres1.CloneSlide(slide, pres2.Slides.LastSlidePosition + 1, pres2, sList);


            //Writing the presentation as a PPT file
            pres2.Write(dataDir + "CloneSlide2.ppt");



        }
    }
}