Imports System
Imports Aspose.Slides.Export
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace VisualBasic.Slides
    Class RemoveHyperlinks
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations()

            ' ExStart:RemoveHyperlinks
            ' Instantiate Presentation class
            Dim presentation As New Presentation(dataDir & Convert.ToString("Hyperlink.pptx"))

            ' Removing the hyperlinks from presentation
            presentation.HyperlinkQueries.RemoveAllHyperlinks()

            ' ExEnd:RemoveHyperlinks
            ' Writing the presentation as a PPTX file
            presentation.Save(dataDir & Convert.ToString("RemovedHyperlink.pptx"), SaveFormat.Pptx)
        End Sub
    End Class
End Namespace