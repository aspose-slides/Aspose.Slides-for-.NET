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
    Public Class AccessOpenDoc
        Public Shared Sub Run()

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            ' Open the ODP file
            Dim pres As New Presentation(dataDir & "AccessOpenDoc.odp")

            ' Saving the ODP presentation to PPTX format
            pres.Save(dataDir & "AccessOpenDoc.pptx", SaveFormat.Pptx)

        End Sub
    End Class
End Namespace