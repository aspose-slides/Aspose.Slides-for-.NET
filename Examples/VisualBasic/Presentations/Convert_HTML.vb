Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports Aspose.Slides.Export

Namespace Aspose.Slides.Examples.VisualBasic.Presentations
    Public Class Convert_HTML
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            ' Instantiate a Presentation object that represents a presentation file
            Using presentation As New Presentation(dataDir & "Convert_HTML.pptx")

                Dim htmlOpt As New HtmlOptions()
                htmlOpt.HtmlFormatter = HtmlFormatter.CreateDocumentFormatter("", False)

                ' Saving the presentation to HTML
                presentation.Save(dataDir & "demo.html", Aspose.Slides.Export.SaveFormat.Html, htmlOpt)
            End Using
        End Sub
    End Class
End Namespace