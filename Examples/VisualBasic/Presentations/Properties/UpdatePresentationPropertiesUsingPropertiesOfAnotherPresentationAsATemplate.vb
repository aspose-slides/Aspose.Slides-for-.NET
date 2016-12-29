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
    Public Class UpdatePresentationPropertiesUsingPropertiesOfAnotherPresentationAsATemplate
        Public Shared Sub Run()
			'ExStart:UpdatePresentationPropertiesUsingPropertiesOfAnotherPresentationAsATemplate
            ' For complete examples and data files, please go to https:// Github.com/aspose-slides/Aspose.Slides-for-.NET

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_PresentationProperties()

            Dim template As DocumentProperties
            Dim info As IPresentationInfo = PresentationFactory.Instance.GetPresentationInfo(dataDir & Convert.ToString("template.pptx"))
            template = DirectCast(info.ReadDocumentProperties(), DocumentProperties)

            template.Author = "Template Author"
            template.Title = "Template Title"
            template.Category = "Template Category"
            template.Keywords = "Keyword1, Keyword2, Keyword3"
            template.Company = "Our Company"
            template.Comments = "Created from template"
            template.ContentType = "Template Content"
            template.Subject = "Template Subject"

            UpdateByTemplate(dataDir & Convert.ToString("doc1.pptx"), template)
            UpdateByTemplate(dataDir & Convert.ToString("doc2.odp"), template)
            UpdateByTemplate(dataDir & Convert.ToString("doc3.ppt"), template)
        End Sub

        Private Shared Sub UpdateByTemplate(path As String, template As IDocumentProperties)
            Dim toUpdate As IPresentationInfo = PresentationFactory.Instance.GetPresentationInfo(path)
            toUpdate.UpdateDocumentProperties(template)
            toUpdate.WriteBindedPresentation(path)
			'ExEnd:UpdatePresentationPropertiesUsingPropertiesOfAnotherPresentationAsATemplate
        End Sub
    End Class
End Namespace
