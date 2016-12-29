using System;
using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Slides.Comments
{
    class AccessSlideComments
    {
        public static void Run()
        {
            //ExStart:AccessSlideComments
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations_Comments();

            // Instantiate Presentation class
            using (Presentation presentation = new Presentation(dataDir + "Comments1.pptx"))
            {
                foreach (var commentAuthor in presentation.CommentAuthors)
                {
                    var author = (CommentAuthor) commentAuthor;
                    foreach (var comment1 in author.Comments)
                    {
                        var comment = (Comment) comment1;
                        // ExEnd:AccessSlideComments
                        Console.WriteLine("ISlide :" + comment.Slide.SlideNumber + " has comment: " + comment.Text + " with Author: " + comment.Author.Name + " posted on time :" + comment.CreatedTime + "\n");
                    }
                }
            }
            //ExEnd:AccessSlideComments
        }
    }
}