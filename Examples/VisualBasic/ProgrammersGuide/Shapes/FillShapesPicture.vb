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
Imports Aspose.Slides.Export
Imports System.Drawing

Namespace VisualBasic.Shapes
    Public Class FillShapesPicture
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not IsExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If

            'Instantiate PrseetationEx class that represents the PPTX
            Using pres As New Presentation()

                'Get the first slide
                Dim sld As ISlide = pres.Slides(0)

                'Add autoshape of rectangle type
                Dim shp As IShape = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 150, 75, 150)


                'Set the fill type to Picture
                shp.FillFormat.FillType = FillType.Picture

                'Set the picture fill mode
                shp.FillFormat.PictureFillFormat.PictureFillMode = PictureFillMode.Tile

                'Set the picture
                Dim img As System.Drawing.Image = CType(New Bitmap(dataDir & "Tulips.jpg"), System.Drawing.Image)
                Dim imgx As IPPImage = pres.Images.AddImage(img)
                shp.FillFormat.PictureFillFormat.Picture.Image = imgx

                'Write the PPTX file to disk
                pres.Save(dataDir & "RectShpPic.pptx", SaveFormat.Pptx)
            End Using
        End Sub
    End Class
End Namespace