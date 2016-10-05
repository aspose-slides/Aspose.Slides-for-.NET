Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides

Namespace Aspose.Slides.Examples.VisualBasic.Presentations
    Public Class SaveAsReadOnly
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

            ' Setting Write protection Password
            pres.ProtectionManager.SetWriteProtection("test")

            ' Save your presentation to a file
            pres.Save(dataDir & "demoWriteProtected_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx)
        End Sub
    End Class
End Namespace