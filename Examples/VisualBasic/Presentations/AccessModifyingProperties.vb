Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports Aspose.Slides.Export

Namespace Aspose.Slides.Examples.VisualBasic.Presentations
    Public Class AccessModifyingProperties
        Public Shared Sub Run()

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            ' Instanciate the Presentation class that represents the PPTX
            Dim pres As New Presentation(dataDir & "AccessModifyingProperties.pptx")

            ' Create a reference to DocumentProperties object associated with Prsentation
            Dim dp As IDocumentProperties = pres.DocumentProperties

            ' Access and modify custom properties
            For i As Integer = 0 To dp.CountOfCustomProperties - 1
                ' Display names and values of custom properties
                System.Console.WriteLine("Custom Property Name : " & dp.GetCustomPropertyName(i))
                System.Console.WriteLine("Custom Property Value : " & dp.GetCustomPropertyName(i))

                ' Modify values of custom properties
                dp(dp.GetCustomPropertyName(i)) = "New Value " & (i + 1)
            Next i

            ' Save your presentation to a file
            pres.Save(dataDir & "CustomDemoModified_out.pptx", SaveFormat.Pptx)

        End Sub
    End Class
End Namespace