Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports System.Drawing

Namespace Aspose.Slides.Examples.VisualBasic.Shapes
    Public Class FormattedRectangle
        Public Shared Sub Run()
			'ExStart:FormattedRectangle		
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Instantiate Prseetation class that represents the PPTX
            Using pres As New Presentation()

                ' Get the first slide
                Dim sld As ISlide = pres.Slides(0)

                ' Add autoshape of rectangle type
                Dim shp As IShape = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 150, 150, 50)

                ' Apply some formatting to rectangle shape
                shp.FillFormat.FillType = FillType.Solid
                shp.FillFormat.SolidFillColor.Color = Color.Chocolate

                ' Apply some formatting to the line of rectangle
                shp.LineFormat.FillFormat.FillType = FillType.Solid
                shp.LineFormat.FillFormat.SolidFillColor.Color = Color.Black
                shp.LineFormat.Width = 5

                ' Write the PPTX file to disk
                pres.Save(dataDir & "RectShp2_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx)

            End Using
			'ExEnd:FormattedRectangle	
        End Sub
    End Class
End Namespace