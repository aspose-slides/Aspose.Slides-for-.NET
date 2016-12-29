using System;
using System.Drawing;
using Aspose.Slides.Export;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Slides.Comments
{
    class AddSlideComments
    {
        public static void Run()
        {
            //ExStart:AddSlideComments
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations_Comments();

            // ExStart:AddSlideComments
            // Instantiate Presentation class
            using (Presentation presentation = new Presentation())
            {
                // Adding Empty slide
                presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);

                // Adding Author
                ICommentAuthor author = presentation.CommentAuthors.AddAuthor("Jawad", "MF");

                // Position of comments
                PointF point = new PointF();
                point.X = 0.2f;
                point.Y = 0.2f;

                // Adding slide comment for an author on slide 1
                author.Comments.AddComment("Hello Jawad, this is slide comment", presentation.Slides[0], point, DateTime.Now);

                // Adding slide comment for an author on slide 1
                author.Comments.AddComment("Hello Jawad, this is second slide comment", presentation.Slides[1], point, DateTime.Now);

                // Accessing ISlide 1
                ISlide slide = presentation.Slides[0];

                // if null is passed as an argument then it will bring comments from all authors on selected slide
                IComment[] Comments = slide.GetSlideComments(author);

                // Accessin the comment at index 0 for slide 1
                String str = Comments[0].Text;

                presentation.Save(dataDir + "Comments_out.pptx", SaveFormat.Pptx);

                if (Comments.GetLength(0) > 0)
                {
                    // ExEnd:AddSlideComments
                    // Select comments collection of Author at index 0
                    ICommentCollection commentCollection = Comments[0].Author.Comments;
                    String Comment = commentCollection[0].Text;
                }
            }
            //ExEnd:AddSlideComments
        }
    }
}