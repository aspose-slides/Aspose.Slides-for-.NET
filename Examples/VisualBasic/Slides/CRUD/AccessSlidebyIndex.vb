Imports System
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Slides.CRUD
    Class AccessSlidebyIndex
        Public Shared Sub Run()
            'ExStart:AccessSlidebyIndex
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations_CRUD()

            ' Create an instance of Presentation class
            Dim presentation As New Presentation(dataDir & Convert.ToString("AccessSlides.pptx"))

            ' Obtain a slide's reference by its index
            Dim slide As ISlide = presentation.Slides(0)
            ' ExEnd:AccessSlidebyIndex
        End Sub
    End Class
End Namespace
