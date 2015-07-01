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
Imports Aspose.Slides.Export

Namespace VisualBasic.Shapes
    Public Class AddSimplePictureFrames
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not IsExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If

            'Instantiate Prseetation class that represents the PPTX
            Using pres As New Presentation()

                'Get the first slide
                Dim sld As ISlide = pres.Slides(0)

                'Instantiate the ImageEx class
                Dim img As System.Drawing.Image = CType(New Bitmap(dataDir & "aspose-logo.jpg"), System.Drawing.Image)
                Dim imgx As IPPImage = pres.Images.AddImage(img)

                'Add Picture Frame with height and width equivalent of Picture
                sld.Shapes.AddPictureFrame(ShapeType.Rectangle, 50, 150, imgx.Width, imgx.Height, imgx)

                'Write the PPTX file to disk
                pres.Save(dataDir & "RectPicFrame.pptx", SaveFormat.Pptx)
            End Using


        End Sub
    End Class
End Namespace