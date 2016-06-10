Imports System.Drawing
Imports Aspose.Slides

'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Namespace VisualBasic.Presentations
    Public Class GetRectangularCoordinatesofParagraph
        Public Shared Sub Run()

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            Using pres As New Presentation(dataDir + "Shapes.pptx")
                Dim shape As IAutoShape = DirectCast(pres.Slides(0).Shapes(0), IAutoShape)
                Dim textFrame = DirectCast(shape.TextFrame, ITextFrame)

                Dim rect As RectangleF = DirectCast(textFrame.Paragraphs(0), Paragraph).GetRect()
            End Using

        End Sub
    End Class
End Namespace