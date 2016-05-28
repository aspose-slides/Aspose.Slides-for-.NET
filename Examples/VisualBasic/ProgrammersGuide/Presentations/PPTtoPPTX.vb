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
Imports Aspose.Slides.Export

Namespace VisualBasic.Presentations
    Public Class PPTtoPPTX
        Public Shared Sub Run()

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            ' Instantiate a Presentation object that represents a PPT file
            Dim presentation As New Presentation(dataDir & "PPTtoPPTX.ppt")

            ' Saving the PPTX presentation to PPTX format
            presentation.Save(dataDir & "PPTtoPPTX.pptx", SaveFormat.Pptx)

        End Sub
    End Class
End Namespace