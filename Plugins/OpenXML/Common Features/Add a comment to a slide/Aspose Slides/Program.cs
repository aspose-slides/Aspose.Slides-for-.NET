// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Slides;
using System;
using System.Drawing;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\..\Sample Files\";
            string FileName = FilePath + "Add a comment to a slide.pptx";
            
            using (Presentation pres = new Presentation())
            {
                //Adding Empty slide
                pres.Slides.AddEmptySlide(pres.LayoutSlides[0]);

                //Adding Autthor
                ICommentAuthor author = pres.CommentAuthors.AddAuthor("Zeeshan", "MZ");

                //Position of comments

                PointF point = new PointF();
                point.X = 1;
                point.Y = 1;

                //Adding slide comment for an author on slide
                author.Comments.AddComment("Hello Zeeshan, this is slide comment", pres.Slides[0], point, DateTime.Now);
                pres.Save(FileName, Aspose.Slides.Export.SaveFormat.Pptx);
            }
        }
    }
}
