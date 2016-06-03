
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Imports System
Imports System.Drawing
Imports Aspose.Slides
Imports Aspose.Slides.Export

Namespace VisualBasic.Presentations
    Public Class GetPositionCoordinatesofPortion
        Public Shared Sub Run()

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            Using presentation As New Presentation(dataDir + "Shapes.pptx")
                Dim shape As IAutoShape = DirectCast(presentation.Slides(0).Shapes(0), IAutoShape)
                Dim textFrame = DirectCast(shape.TextFrame, ITextFrame)

                For Each paragraph As Paragraph In textFrame.Paragraphs
                    For Each portion As Portion In paragraph.Portions
                        Dim point As PointF = portion.GetCoordinates()
                        Console.Write("{0}Corrdinates X ={1} Corrdinates Y ={2}", Environment.NewLine, point.X, point.Y)
                    Next
                Next
            End Using
        End Sub
    End Class
End Namespace

