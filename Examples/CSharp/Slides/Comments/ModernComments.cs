using System;
using System.Drawing;
using System.IO;
using Aspose.Slides.Export;

/*
This example demonstrates the addition of a modern comment to a slide
*/

namespace Aspose.Slides.Examples.CSharp.Slides.Comments
{
    class ModernComments
    {
        public static void Run()
        {
            // The path to the output file.
            string outPptxFile = Path.Combine(RunExamples.OutPath, "ModernComments_out.pptx");

            using (Presentation pres = new Presentation())
            {
                // Add author
                ICommentAuthor newAuthor = pres.CommentAuthors.AddAuthor("Some Author", "SA");

                // Add comment
                IModernComment modernComment = newAuthor.Comments.AddModernComment("This is a modern comment", pres.Slides[0], null, new PointF(100, 100), DateTime.Now);

                // Save presentation
                pres.Save(outPptxFile, SaveFormat.Pptx);
            }
        }
    }
}