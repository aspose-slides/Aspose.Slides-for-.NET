Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports Aspose.Slides.Export
Imports System.Drawing

Namespace Aspose.Slides.Examples.VisualBasic.Shapes
    Public Class FillShapesPicture
        Public Shared Sub Run()
		  'ExStart:FillShapesPicture		
		  ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Instantiate PrseetationEx class that represents the PPTX
            Using presentation As New Presentation()

                ' Get the first slide
                Dim islide As ISlide = presentation.Slides(0)

                ' Add autoshape of rectangle type
                Dim iShape As IShape = islide.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 150, 75, 150)

                ' Set the fill type to Picture
                iShape.FillFormat.FillType = FillType.Picture

                ' Set the picture fill mode
                iShape.FillFormat.PictureFillFormat.PictureFillMode = PictureFillMode.Tile

                ' Set the picture
                Dim img As System.Drawing.Image = CType(New Bitmap(dataDir & "Tulips.jpg"), System.Drawing.Image)
                Dim imgx As IPPImage = presentation.Images.AddImage(img)
                iShape.FillFormat.PictureFillFormat.Picture.Image = imgx

                ' Write the PPTX file to disk
                presentation.Save(dataDir & "RectShpPic_out.pptx", SaveFormat.Pptx)

            End Using
			'ExStart:FillShapesPicture		
        End Sub
    End Class
End Namespace