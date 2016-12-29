Imports System
Imports System.Drawing
Imports Aspose.Slides.Export

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Slides.Comments
    Class AddSlideComments
        Public Shared Sub Run()
            ' ExStart:AddSlideComments
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations_Comments()

            ' Instantiate Presentation class
            Using presentation As New Presentation()
                ' Adding Empty slide
                presentation.Slides.AddEmptySlide(presentation.LayoutSlides(0))

                ' Adding Author
                Dim author As ICommentAuthor = presentation.CommentAuthors.AddAuthor("Jawad", "MF")

                ' Position of comments
                Dim point As New PointF()
                point.X = 0.2F
                point.Y = 0.2F

                ' Adding slide comment for an author on slide 1
                author.Comments.AddComment("Hello Jawad, this is slide comment", presentation.Slides(0), point, DateTime.Now)

                ' Adding slide comment for an author on slide 1
                author.Comments.AddComment("Hello Jawad, this is second slide comment", presentation.Slides(1), point, DateTime.Now)

                ' Accessing ISlide 1
                Dim slide As ISlide = presentation.Slides(0)

                ' if null is passed as an argument then it will bring comments from all authors on selected slide
                Dim Comments As IComment() = slide.GetSlideComments(author)

                ' Accessin the comment at index 0 for slide 1
                Dim str As [String] = Comments(0).Text

                presentation.Save(dataDir & Convert.ToString("Comments_out.pptx"), SaveFormat.Pptx)

                If Comments.GetLength(0) > 0 Then
                    ' Select comments collection of Author at index 0
                    Dim commentCollection As ICommentCollection = Comments(0).Author.Comments
                    Dim Comment As [String] = commentCollection(0).Text
                End If
            End Using
            ' ExEnd:AddSlideComments
        End Sub
    End Class
End Namespace
