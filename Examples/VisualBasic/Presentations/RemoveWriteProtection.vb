'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Slides. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides

Namespace VisualBasic.Presentations
    Public Class RemoveWriteProtection
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            'Opening the presentation file
            Dim presentation As New Presentation(dataDir & "RemoveWriteProtection.pptx")


            'Checking if presentation is write protected
            If presentation.ProtectionManager.IsWriteProtected Then
                'Removing Write protection
                presentation.ProtectionManager.RemoveWriteProtection()
            End If

            'Saving presentation
            presentation.Save(dataDir & "newDemo.pptx", Aspose.Slides.Export.SaveFormat.Pptx)
        End Sub
    End Class
End Namespace