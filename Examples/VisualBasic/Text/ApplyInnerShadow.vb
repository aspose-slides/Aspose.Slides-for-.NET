Imports System
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Text
    Class ApplyInnerShadow
        Public Shared Sub Run()
            ' ExStart:ApplyInnerShadow

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Text()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If Not IsExists Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If

            ' Instantiate PresentationEx 
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
                pres.Save(dataDir & Convert.ToString("ApplyInnerShadow_out.pptx"), Aspose.Slides.Export.SaveFormat.Pptx)
            End Using
            ' ExStart:ApplyInnerShadow
        End Sub
    End Class
End Namespace