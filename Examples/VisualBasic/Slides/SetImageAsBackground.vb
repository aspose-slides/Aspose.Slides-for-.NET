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
Imports System.Drawing
Imports Aspose.Slides.Export

Namespace VisualBasic.Slides
    Public Class SetImageAsBackground
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations()

            'Instantiate the Presentation class that represents the presentation file

            Using pres As New Presentation(dataDir & "SetImageAsBackground.pptx")

                'Set the background with Image
                pres.Slides(0).Background.Type = BackgroundType.OwnBackground
                pres.Slides(0).Background.FillFormat.FillType = FillType.Picture
                pres.Slides(0).Background.FillFormat.PictureFillFormat.PictureFillMode = PictureFillMode.Stretch

                'Set the picture
                Dim img As System.Drawing.Image = CType(New Bitmap(dataDir & "Tulips.jpg"), System.Drawing.Image)

                'Add image to presentation's images collection
                Dim imgx As IPPImage = pres.Images.AddImage(img)

                pres.Slides(0).Background.FillFormat.PictureFillFormat.Picture.Image = imgx

                'Write the presentation to disk
                pres.Save(dataDir & "ContentBG_Img.pptx", SaveFormat.Pptx)

            End Using
        End Sub
    End Class
End Namespace