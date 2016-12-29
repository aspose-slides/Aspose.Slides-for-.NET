Imports System
Imports Aspose.Slides.Export

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Presentations
    Public Class ModifyBuiltinProperties
        Public Shared Sub Run()
			'ExStart:ModifyBuiltinProperties
            ' For complete examples and data files, please go to https:// Github.com/aspose-slides/Aspose.Slides-for-.NET

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_PresentationProperties()

            ' Instantiate the Presentation class that represents the Presentation
            Dim presentation As New Presentation(dataDir & Convert.ToString("ModifyBuiltinProperties.pptx"))

            ' Create a reference to IDocumentProperties object associated with Presentation
            Dim documentProperties As IDocumentProperties = presentation.DocumentProperties

            ' Set the builtin properties
            documentProperties.Author = "Aspose.Slides for .NET"
            documentProperties.Title = "Modifying Presentation Properties"
            documentProperties.Subject = "Aspose Subject"
            documentProperties.Comments = "Aspose Description"
            documentProperties.Manager = "Aspose Manager"

            ' Save your presentation to a file
            presentation.Save(dataDir & Convert.ToString("DocumentProperties_out.pptx"), SaveFormat.Pptx)

			'ExEnd:ModifyBuiltinProperties
		End Sub
    End Class
End Namespace
