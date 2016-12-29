Imports System
Imports Aspose.Slides.Export

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Presentations.Conversion
    Class CreateNewPresentation
        Public Shared Sub Run()
			'ExStart:CreateNewPresentation
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Conversion()

            ' Instantiate a Presentation object that represents a presentation file
            Using presentation As New Presentation()
                ' Get the first slide
                Dim slide As ISlide = presentation.Slides(0)

                ' Add an autoshape of type line
                slide.Shapes.AddAutoShape(ShapeType.Line, 50, 150, 300, 0)
                presentation.Save(dataDir & Convert.ToString("NewPresentation_out.pptx"), SaveFormat.Pptx)
            End Using
			'ExEnd:CreateNewPresentation
        End Sub
    End Class
End Namespace
