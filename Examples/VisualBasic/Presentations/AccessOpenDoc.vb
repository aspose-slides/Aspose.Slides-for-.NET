Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports Aspose.Slides.Export

Namespace Aspose.Slides.Examples.VisualBasic.Presentations
    Public Class AccessOpenDoc
        Public Shared Sub Run()

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            ' Open the ODP file
            Dim pres As New Presentation(dataDir & "AccessOpenDoc.odp")

            ' Saving the ODP presentation to PPTX format
            pres.Save(dataDir & "AccessOpenDoc_out.pptx", SaveFormat.Pptx)

        End Sub
    End Class
End Namespace