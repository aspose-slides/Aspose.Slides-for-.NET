Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides

Namespace Aspose.Slides.Examples.VisualBasic.Text
    Public Class TextBoxOnSlideProgram
        Public Shared Sub Run()
            ' ExStart:TextBoxOnSlideProgram
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Text()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not IsExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If

            ' Instantiate PresentationEx//Instantiate PresentationEx
            Using pres As New Presentation()

                ' Get the first slide
                Dim sld As ISlide = pres.Slides(0)

                ' Add an AutoShape of Rectangle type
                Dim ashp As IAutoShape = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 150, 75, 150, 50)

                ' Add TextFrame to the Rectangle
                ashp.AddTextFrame(" ")

                ' Accessing the text frame
                Dim txtFrame As ITextFrame = ashp.TextFrame

                ' Create the Paragraph object for text frame
                Dim para As IParagraph = txtFrame.Paragraphs(0)

                ' Create Portion object for paragraph
                Dim portion As IPortion = para.Portions(0)

                ' Set Text
                portion.Text = "Aspose TextBox"

                ' Save the presentation to disk
                pres.Save(dataDir & "TextBox_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx)
            End Using
            ' ExEnd:TextBoxOnSlideProgram
        End Sub
    End Class
End Namespace