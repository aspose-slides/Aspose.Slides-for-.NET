// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using System;
using System.Drawing;
using Aspose.Slides.Pptx;

namespace Aspose_Slides
{
    class Program
    {
        static void Main(string[] args)
        {

            using (PresentationEx pres = new PresentationEx())
            {
                //Adding Empty slide
                pres.Slides.AddEmptySlide(pres.LayoutSlides[0]);

                //Adding Autthor
                CommentAuthorEx author = pres.CommentAuthors.AddAuthor("Zeeshan", "MZ");

                //Position of comments

                PointF point = new PointF();
                point.X = 1;
                point.Y = 1;

                //Adding slide comment for an author on slide
                author.Comments.AddComment("Hello Zeeshan, this is slide comment", pres.Slides[0], point, DateTime.Now);
                pres.Write("Comments.pptx");
            }
        }
    }
}
