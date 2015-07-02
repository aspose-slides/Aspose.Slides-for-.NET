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
Imports System.Drawing

Namespace VisualBasic.Slides
    Public Class ThumbnailFromSlide
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations()

            'Instantiate a Presentation class that represents the presentation file
            Using pres As New Presentation(dataDir & "ThumbnailFromSlide.pptx")

                'Access the first slide
                Dim sld As ISlide = pres.Slides(0)

                'Create a full scale image
                Dim bmp As Bitmap = sld.GetThumbnail(1.0F, 1.0F)

                'Save the image to disk in JPEG format
                bmp.Save(dataDir & "Thumbnail.jpg", System.Drawing.Imaging.ImageFormat.Jpeg)

            End Using
        End Sub
    End Class
End Namespace