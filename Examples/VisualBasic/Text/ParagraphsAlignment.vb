Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports Aspose.Slides.Export

Namespace Aspose.Slides.Examples.VisualBasic.Text
    Public Class ParagraphsAlignment
        Public Shared Sub Run()
            ' ExStart:ParagraphsAlignment

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Text()

            ' Instantiate a Presentation object that represents a PPTX file
            Using pres As New Presentation(dataDir & "ParagraphsAlignment.pptx")

                ' Accessing first slide
                Dim slide As ISlide = pres.Slides(0)

                ' Accessing the first and second placeholder in the slide and typecasting it as AutoShape
                Dim tf1 As ITextFrame = (CType(slide.Shapes(0), IAutoShape)).TextFrame
                Dim tf2 As ITextFrame = (CType(slide.Shapes(1), IAutoShape)).TextFrame

                ' Change the text in both placeholders
                tf1.Text = "Center Align by Aspose"
                tf2.Text = "Center Align by Aspose"

                ' Getting the first paragraph of the placeholders
                Dim para1 As IParagraph = tf1.Paragraphs(0)
                Dim para2 As IParagraph = tf2.Paragraphs(0)

                ' Aligning the text paragraph to center
                para1.ParagraphFormat.Alignment = TextAlignment.Center
                para2.ParagraphFormat.Alignment = TextAlignment.Center

                'Writing the presentation as a PPTX file
                pres.Save(dataDir & "Centeralign_out.pptx", SaveFormat.Pptx)
            End Using
            ' ExStart:ParagraphsAlignment
        End Sub
    End Class
End Namespace