Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides

Namespace Aspose.Slides.Examples.VisualBasic.Presentations
    Public Class RemoveWriteProtection
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            ' Opening the presentation file
            Dim presentation As New Presentation(dataDir & "RemoveWriteProtection.pptx")


            ' Checking if presentation is write protected
            If presentation.ProtectionManager.IsWriteProtected Then
                ' Removing Write protection
                presentation.ProtectionManager.RemoveWriteProtection()
            End If

            ' Saving presentation
            presentation.Save(dataDir & "newDemo_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx)
        End Sub
    End Class
End Namespace