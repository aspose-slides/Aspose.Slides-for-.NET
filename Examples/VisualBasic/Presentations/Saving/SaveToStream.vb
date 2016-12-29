Imports System
Imports System.IO
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Presentations.Saving
    Class SaveToStream
        Public Shared Sub Run()
			'ExStart:SaveToStream
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_PresentationSaving()

            ' Instantiate a Presentation object that represents a PPT file
            Using presentation As New Presentation()

                Dim shape As IAutoShape = presentation.Slides(0).Shapes.AddAutoShape(ShapeType.Rectangle, 200, 200, 200, 200)

                ' Add text to shape
                shape.TextFrame.Text = "This demo shows how to Create PowerPoint file and save it to Stream."

                Dim toStream As New FileStream(dataDir & Convert.ToString("Save_As_Stream_out.pptx"), FileMode.Create)
                presentation.Save(toStream, Aspose.Slides.Export.SaveFormat.Pptx)
                toStream.Close()
            End Using
			'ExEnd:SaveToStream
        End Sub
    End Class
End Namespace
