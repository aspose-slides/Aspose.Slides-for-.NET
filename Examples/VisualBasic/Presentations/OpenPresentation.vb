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
    Public Class OpenPresentation
        Public Shared Sub Run()
            ' For complete examples and data files, please go to https://github.com/aspose-slides/Aspose.Slides-for-.NET

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            ' Opening the presentation file by passing the file path to the constructor of Presentation class
            Dim presentation As New Presentation(dataDir & "OpenPresentation.pptx")

            ' Printing the total number of slides present in the presentation
            System.Console.WriteLine(presentation.Slides.Count.ToString())

        End Sub
    End Class
End Namespace