Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports System.Drawing.Imaging
Imports Aspose.Slides.Export

Namespace Aspose.Slides.Examples.VisualBasic.Text
    Public Class DefaultFonts
        Public Shared Sub Run()
            ' ExStart:FontFamily

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Text()

            ' Use load options to define the default regualr and asian fonts//Use load options to define the default regualr and asian fonts
            Dim lo As New LoadOptions(LoadFormat.Auto)
            lo.DefaultRegularFont = "Wingdings"
            lo.DefaultAsianFont = "Wingdings"

            ' Load the presentation
            Using pptx As New Presentation(dataDir & "DefaultFonts.pptx", lo)

                ' Generate slide thumbnail
                pptx.Slides(0).GetThumbnail(1, 1).Save(dataDir & "DefaultFonts_png_out.png", ImageFormat.Png)

                ' Generate PDF
                pptx.Save(dataDir & "DefaultFonts_pdf_out.pdf", SaveFormat.Pdf)

                ' Generate XPS
                pptx.Save(dataDir & "DefaultFonts_xps_out.xps", SaveFormat.Xps)
            End Using
            ' ExEnd:FontFamily
        End Sub
    End Class
End Namespace