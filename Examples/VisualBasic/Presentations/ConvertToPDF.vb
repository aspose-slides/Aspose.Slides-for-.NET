Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides

Namespace Aspose.Slides.Examples.VisualBasic.Presentations
    Public Class ConvertToPDF
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            'Instantiate a Presentation object that represents a presentation file
            Dim pres As New Presentation(dataDir & "ConvertToPDF.pptx")

            'Save the presentation to PDF with default options
            pres.Save(dataDir & "PDFUsingDefaultOptions.pdf", Aspose.Slides.Export.SaveFormat.Pdf)

        End Sub
    End Class
End Namespace