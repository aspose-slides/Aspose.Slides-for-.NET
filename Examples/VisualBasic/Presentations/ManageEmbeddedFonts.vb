Imports System
Imports System.Drawing
Imports System.Drawing.Imaging
Imports Aspose.Slides.Examples.VisualBasic
Imports Aspose.Slides.Export

Namespace Aspose.Slides.Examples.VisualBasic.Presentations
    Public Class ManageEmbeddedFonts
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            ' Instantiate a Presentation object that represents a presentation file
            Using presentation As New Presentation(dataDir & Convert.ToString("EmbeddedFonts.pptx"))
                ' render a slide that contains a text frame that uses embedded "FunSized"
                presentation.Slides(0).GetThumbnail(New Size(960, 720)).Save(dataDir & Convert.ToString("picture1_out.png"), ImageFormat.Png)

                Dim fontsManager As IFontsManager = presentation.FontsManager

                ' get all embedded fonts
                Dim embeddedFonts As IFontData() = fontsManager.GetEmbeddedFonts()

                ' find "FunSized" font
                Dim funSizedEmbeddedFont As IFontData = Array.Find(embeddedFonts, Function(data As IFontData) data.FontName = "Calibri")

                ' remove "Calibri" font
                fontsManager.RemoveEmbeddedFont(funSizedEmbeddedFont)

                ' render the presentation; removed "Calibri" font is replaced to an existing one
                presentation.Slides(0).GetThumbnail(New Size(960, 720)).Save(dataDir & Convert.ToString("picture2_out.png"), ImageFormat.Png)

                ' save the presentation without embedded "Calibri" font
                presentation.Save(dataDir & Convert.ToString("WithoutManageEmbeddedFonts_out.ppt"), SaveFormat.Ppt)
            End Using
        End Sub
    End Class
End Namespace
