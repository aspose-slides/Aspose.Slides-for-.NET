Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports System.Text

Namespace Aspose.Slides.Examples.VisualBasic.Text
    Public Class ExportingHTMLText
        Public Shared Sub Run()
            ' ExStart:FontFamily
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Text()
            ' Load the presentation file
            Using pres As New Presentation(dataDir & "ExportingHTMLText.pptx")

                ' Acesss the default first slide of presentation
                Dim slide As ISlide = pres.Slides(0)

                ' Desired index
                Dim index As Integer = 0

                ' Accessing the added shape
                Dim ashape As IAutoShape = CType(slide.Shapes(index), IAutoShape)

                ' Extracting first paragraph as HTML
                Dim sw As New StreamWriter(dataDir & "ExportingHTMLText_out.html", False, Encoding.UTF8)

                'Writing Paragraphs data to HTML by providing paragraph starting index, total paragraphs to be copied
                sw.Write(ashape.TextFrame.Paragraphs.ExportToHtml(0, ashape.TextFrame.Paragraphs.Count, Nothing))

                sw.Close()
            End Using
            ' ExEnd:FontFamily

        End Sub
    End Class
End Namespace