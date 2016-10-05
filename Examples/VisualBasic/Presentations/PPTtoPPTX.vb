Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports Aspose.Slides.Export

Namespace Aspose.Slides.Examples.VisualBasic.Presentations
    Public Class PPTtoPPTX
        Public Shared Sub Run()

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            ' Instantiate a Presentation object that represents a PPT file
            Dim presentation As New Presentation(dataDir & "PPTtoPPTX.ppt")

            ' Saving the PPTX presentation to PPTX format
            presentation.Save(dataDir & "PPTtoPPTX_out.pptx", SaveFormat.Pptx)

        End Sub
    End Class
End Namespace