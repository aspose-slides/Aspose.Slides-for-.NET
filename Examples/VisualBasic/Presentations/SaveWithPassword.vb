Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides

Namespace Aspose.Slides.Examples.VisualBasic.Presentations
    Public Class SaveWithPassword
        Public Shared Sub Run()

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            ' Create directory if it is not already present.
            Dim isExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not isExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If

            ' Instantiate a Presentation object that represents a PPT file
            Dim presentation As New Presentation()

            ' ....do some work here.....

            ' Setting Password
            presentation.ProtectionManager.Encrypt("pass")

            ' Save your presentation to a file
            presentation.Save(dataDir & "Saving_Password_Protected_Presentation.pptx", Export.SaveFormat.Pptx)

        End Sub
    End Class
End Namespace