Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides

Namespace Aspose.Slides.Examples.VisualBasic.Presentations
    Public Class SaveProperties
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not IsExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If

            ' Instantiate a Presentation object that represents a PPT file
            Dim pres As New Presentation()

            '....do some work here.....

            ' Setting access to document properties in password protected mode
            pres.ProtectionManager.EncryptDocumentProperties = False

            ' Setting Password
            pres.ProtectionManager.Encrypt("pass")

            ' Save your presentation to a file
            pres.Save(dataDir & "demoPassDocument_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx)
        End Sub
    End Class
End Namespace