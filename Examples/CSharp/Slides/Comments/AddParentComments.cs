using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Slides.Comments
{
    class AddParentComments
    {
        public static void Run() {

            //ExStart:AddParentComments
            // The path to the output directory.
            string outPptxFile = RunExamples.OutPath;

            using (Presentation pres = new Presentation())
            {
                // Add comment
                ICommentAuthor author1 = pres.CommentAuthors.AddAuthor("Author_1", "A.A.");
                IComment comment1 = author1.Comments.AddComment("comment1", pres.Slides[0], new PointF(10, 10), DateTime.Now);

                // Add reply for comment1
                ICommentAuthor author2 = pres.CommentAuthors.AddAuthor("Autror_2", "B.B.");
                IComment reply1 = author2.Comments.AddComment("reply 1 for comment 1", pres.Slides[0], new PointF(10, 10), DateTime.Now);
                reply1.ParentComment = comment1;

                // Add reply for comment1
                IComment reply2 = author2.Comments.AddComment("reply 2 for comment 1", pres.Slides[0], new PointF(10, 10), DateTime.Now);
                reply2.ParentComment = comment1;

                // Add reply to reply
                IComment subReply = author1.Comments.AddComment("subreply 3 for reply 2", pres.Slides[0], new PointF(10, 10), DateTime.Now);
                subReply.ParentComment = reply2;

                IComment comment2 = author2.Comments.AddComment("comment 2", pres.Slides[0], new PointF(10, 10), DateTime.Now);
                IComment comment3 = author2.Comments.AddComment("comment 3", pres.Slides[0], new PointF(10, 10), DateTime.Now);

                IComment reply3 = author1.Comments.AddComment("reply 4 for comment 3", pres.Slides[0], new PointF(10, 10), DateTime.Now);
                reply3.ParentComment = comment3;

                // Display hierarchy on console
                ISlide slide = pres.Slides[0];
                var comments = slide.GetSlideComments(null);
                for (int i = 0; i < comments.Length; i++)
                {
                    IComment comment = comments[i];
                    while (comment.ParentComment != null)
                    {
                        Console.Write("\t");
                        comment = comment.ParentComment;
                    }

                    Console.Write("{0} : {1}", comments[i].Author.Name, comments[i].Text);
                    Console.WriteLine();
                }

                pres.Save(outPptxFile + "parent_comment.pptx", SaveFormat.Pptx);

                // Remove comment1 and all its replies
                comment1.Remove();

                pres.Save(outPptxFile + "remove_comment.pptx", SaveFormat.Pptx);
            }
            //ExEnd:AddParentComments
        }
    }
}
