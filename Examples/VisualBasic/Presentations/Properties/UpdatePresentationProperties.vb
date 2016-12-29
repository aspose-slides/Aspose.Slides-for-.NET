Imports System
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Presentations.Properties
    Public Class UpdatePresentationProperties
        Public Shared Sub Run()
			'ExStart:UpdatePresentationProperties
            ' For complete examples and data files, please go to https:// Github.com/aspose-slides/Aspose.Slides-for-.NET

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_PresentationProperties()

            ' read the info of presentation 
            Dim info As IPresentationInfo = PresentationFactory.Instance.GetPresentationInfo(dataDir & Convert.ToString("ModifyBuiltinProperties1.pptx"))

            ' obtain the current properties 
            Dim props As IDocumentProperties = info.ReadDocumentProperties()

            ' set the new values of Author and Title fields 
            props.Author = "New Author"
            props.Title = "New Title"

            ' update the presentation with a new values 
            info.UpdateDocumentProperties(props)
            info.WriteBindedPresentation(dataDir & Convert.ToString("ModifyBuiltinProperties1.pptx"))
			
			'ExEnd:UpdatePresentationProperties
        End Sub
    End Class
End Namespace
