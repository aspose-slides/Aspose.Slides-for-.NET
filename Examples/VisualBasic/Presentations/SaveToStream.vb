Imports System
Imports System.Drawing
Imports System.IO
Imports Aspose.Slides

'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Namespace Aspose.Slides.Examples.VisualBasic.Presentations
    Public Class SaveToStream
        Public Shared Sub Run()

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            ' Instantiate a Presentation object that represents a PPT file
            Using presentation As New Presentation()

                ' add shape in presentation
                Dim shape As IAutoShape = presentation.Slides(0).Shapes.AddAutoShape(ShapeType.Rectangle, 200, 200, 200, 200)

                ' add text to shape
                shape.TextFrame.Text = "This demo shows how to Create PowerPoint file and save it to Stream."

                Dim toStream As New FileStream(dataDir & Convert.ToString("Save_As_Stream_out.pptx"), FileMode.Create)
                presentation.Save(toStream, Export.SaveFormat.Pptx)
                toStream.Close()

            End Using
        End Sub
    End Class
End Namespace