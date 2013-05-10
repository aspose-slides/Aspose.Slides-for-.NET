//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;
using System.Drawing;
using System;

namespace AddingSlideComments
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            Presentation pres = new Presentation();

            //Getting first slide
            Slide slide = pres.Slides[0];

            //Adding Autthor
            CommentAuthor author = pres.CommentAuthors.AddAuthor("Aspose");


            //Position of comments
            Point point = new Point();
            point.X = 100;
            point.Y = 100;

            //Adding Slide comments
            slide.SlideComments.AddComment(author, "AP", "Hello Aspose, this is slide comment", DateTime.Now, point);

            //Adding Empty slide
            slide = pres.AddEmptySlide();

            //Position of comments
            Point point2 = new Point();
            point2.X = 500;
            point2.Y = 1400;

            //Adding Slide comments
            slide.SlideComments.AddComment(author, "AP", "Hello Aspose, this is second slide comment", DateTime.Now, point2);

            CommentCollection comments = slide.SlideComments;
            //Accessin the comment at index 0 for slide 1
            String str = comments[0].Text;

            pres.Write(dataDir + "Comments.ppt");
        }
    }
}