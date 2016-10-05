

Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides

Namespace Aspose.Slides.Examples.VisualBasic.Presentations
    Public Class Convert_XPS_Options
        Public Shared Sub Run()

            ' For complete examples and data files, please go to https://github.com/aspose-slides/Aspose.Slides-for-.NET

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            ' Instantiate a Presentation object that represents a presentation file
            Using presentation As New Presentation(dataDir & "Convert_XPS_Options.pptx")

                ' Instantiate the TiffOptions class
                Dim options As New Aspose.Slides.Export.XpsOptions()

                ' Save MetaFiles as PNG
                options.SaveMetafilesAsPng = True

                ' Save the presentation to XPS document
                presentation.Save(dataDir & "XPS_With_Options_out.xps", Aspose.Slides.Export.SaveFormat.Xps, options)

            End Using
        End Sub
    End Class
End Namespace