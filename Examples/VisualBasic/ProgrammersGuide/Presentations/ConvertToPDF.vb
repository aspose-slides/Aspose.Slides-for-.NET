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
    Public Class ConvertToPDF
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            'Instantiate a Presentation object that represents a presentation file
            Dim pres As New Presentation(dataDir & "ConvertToPDF.pptx")

            'Save the presentation to PDF with default options
            pres.Save(dataDir & "PDFUsingDefaultOptions.pdf", Aspose.Slides.Export.SaveFormat.Pdf)

        End Sub
    End Class
End Namespace