Imports System
Imports Aspose.Slides.Export

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Slides.Hyperlink
    Class RemoveHyperlinks
        Public Shared Sub Run()
            ' ExStart:RemoveHyperlinks
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations_Hyperlink()

            ' Instantiate Presentation class
            Dim presentation As New Presentation(dataDir & Convert.ToString("Hyperlink.pptx"))

            ' Removing the hyperlinks from presentation
            presentation.HyperlinkQueries.RemoveAllHyperlinks()

            ' Writing the presentation as a PPTX file
            presentation.Save(dataDir & Convert.ToString("RemovedHyperlink_out.pptx"), SaveFormat.Pptx)
            ' ExEnd:RemoveHyperlinks
        End Sub
    End Class
End Namespace
