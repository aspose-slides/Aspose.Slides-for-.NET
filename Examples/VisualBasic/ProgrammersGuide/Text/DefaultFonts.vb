'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Slides. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports System.Drawing.Imaging
Imports Aspose.Slides.Export

Namespace VisualBasic.Text
    Public Class DefaultFonts
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Text()

            'Use load options to define the default regualr and asian fonts//Use load options to define the default regualr and asian fonts
            Dim lo As New LoadOptions(LoadFormat.Auto)
            lo.DefaultRegularFont = "Wingdings"
            lo.DefaultAsianFont = "Wingdings"

            'Load the presentation
            Using pptx As New Presentation(dataDir & "DefaultFonts.pptx", lo)

                'Generate slide thumbnail
                pptx.Slides(0).GetThumbnail(1, 1).Save(dataDir & "output.png", ImageFormat.Png)

                'Generate PDF
                pptx.Save("output.pdf", SaveFormat.Pdf)

                'Generate XPS
                pptx.Save(dataDir & "output.xps", SaveFormat.Xps)
            End Using



        End Sub
    End Class
End Namespace