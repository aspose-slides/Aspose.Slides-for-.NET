Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides

Namespace Aspose.Slides.Examples.VisualBasic.Presentations
    Public Class ModifyBuiltinProperties
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            ' Instantiate the Presentation class that represents the Presentation
            Dim pres As New Presentation(dataDir & "ModifyBuiltinProperties.pptx")

            ' Create a reference to IDocumentProperties object associated with Presentation
            Dim dp As IDocumentProperties = pres.DocumentProperties

            ' Set the builtin properties
            dp.Author = "Aspose.Slides for .NET"
            dp.Title = "Modifying Presentation Properties"
            dp.Subject = "Aspose Subject"
            dp.Comments = "Aspose Description"
            dp.Manager = "Aspose Manager"

            ' Save your presentation to a file
            pres.Save(dataDir & "Updated_Document_Properties_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx)

        End Sub
    End Class
End Namespace