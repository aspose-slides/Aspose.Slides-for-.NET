Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports System.Drawing
Imports Aspose.Slides.Export

Namespace Aspose.Slides.Examples.VisualBasic.Shapes
    Public Class FormatJoinStyles
        Public Shared Sub Run()
			'ExStart:FormatJoinStyles	
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not IsExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If

            ' Instantiate Prseetation class that represents the PPTX
            Using pres As New Presentation()

                ' Get the first slide
                Dim sld As ISlide = pres.Slides(0)

                ' Add three autoshapes of rectangle type
                Dim shp1 As IShape = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 100, 150, 75)

                Dim shp2 As IShape = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 300, 100, 150, 75)

                Dim shp3 As IShape = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 250, 150, 75)

                ' Set the fill color of the rectangle shape
                shp1.FillFormat.FillType = FillType.Solid
                shp1.FillFormat.SolidFillColor.Color = Color.Black
                shp2.FillFormat.FillType = FillType.Solid
                shp2.FillFormat.SolidFillColor.Color = Color.Black
                shp3.FillFormat.FillType = FillType.Solid
                shp3.FillFormat.SolidFillColor.Color = Color.Black

                ' Set the line width
                shp1.LineFormat.Width = 15
                shp2.LineFormat.Width = 15
                shp3.LineFormat.Width = 15

                ' Set the color of the line of rectangle
                shp1.LineFormat.FillFormat.FillType = FillType.Solid
                shp1.LineFormat.FillFormat.SolidFillColor.Color = Color.Blue
                shp2.LineFormat.FillFormat.FillType = FillType.Solid
                shp2.LineFormat.FillFormat.SolidFillColor.Color = Color.Blue
                shp3.LineFormat.FillFormat.FillType = FillType.Solid
                shp3.LineFormat.FillFormat.SolidFillColor.Color = Color.Blue

                ' Set the Join Style
                shp1.LineFormat.JoinStyle = LineJoinStyle.Miter
                shp2.LineFormat.JoinStyle = LineJoinStyle.Bevel
                shp3.LineFormat.JoinStyle = LineJoinStyle.Round

                ' Add text to each rectangle
                CType(shp1, IAutoShape).TextFrame.Text = "This is Miter Join Style"
                CType(shp2, IAutoShape).TextFrame.Text = "This is Bevel Join Style"
                CType(shp3, IAutoShape).TextFrame.Text = "This is Round Join Style"

                'Write the PPTX file to disk
                pres.Save(dataDir & "RectShpLnJoin_out.pptx", SaveFormat.Pptx)
            End Using
			'ExEnd:FormatJoinStyles	
		End Sub
    End Class
End Namespace