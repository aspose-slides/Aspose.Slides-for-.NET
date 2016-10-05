
Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides

Namespace Aspose.Slides.Examples.VisualBasic.Presentations
    Public Class Convert_XPS
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            ' Instantiate a Presentation object that represents a presentation file
            Using pres As New Presentation(dataDir & "Convert_XPS.pptx")

                ' Saving the presentation to TIFF document
                pres.Save(dataDir & "XPS_Output_Without_XPSOption_out.xps", Export.SaveFormat.Xps)
            End Using
        End Sub
    End Class
End Namespace