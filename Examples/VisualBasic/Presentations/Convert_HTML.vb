Imports System
Imports Aspose.Slides.Export

Namespace Aspose.Slides.Examples.VisualBasic.Presentations
    Public Class Convert_HTML
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            ' Instantiate a Presentation object that represents a presentation file
            Using presentation As New Presentation(dataDir & Convert.ToString("Convert_HTML.pptx"))
                Dim controller As New ResponsiveHtmlController()
                Dim htmlOptions As New HtmlOptions() With { _
                    .HtmlFormatter = HtmlFormatter.CreateCustomFormatter(controller) _
                }

                ' Saving the presentation to HTML
                presentation.Save(dataDir & Convert.ToString("Convert_HTML_out.html"), Aspose.Slides.Export.SaveFormat.Html, htmlOptions)
            End Using
        End Sub
    End Class
End Namespace