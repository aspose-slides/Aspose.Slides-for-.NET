Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports Aspose.Slides.Export
Imports System.Drawing

Namespace Aspose.Slides.Examples.VisualBasic.Shapes
    Public Class FormattedEllipse
        Public Shared Sub Run()

			'ExStart:FormattedEllipse	
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Instantiate Prseetation class that represents the PPTX
            Using presentation As New Presentation()

                ' Get the first slide
                Dim islide As ISlide = presentation.Slides(0)

                ' Add autoshape of ellipse type
                Dim ishape As IShape = islide.Shapes.AddAutoShape(ShapeType.Ellipse, 50, 150, 150, 50)

                ' Apply some formatting to ellipse shape
                ishape.FillFormat.FillType = FillType.Solid
                ishape.FillFormat.SolidFillColor.Color = Color.Chocolate

                ' Apply some formatting to the line of Ellipse
                ishape.LineFormat.FillFormat.FillType = FillType.Solid
                ishape.LineFormat.FillFormat.SolidFillColor.Color = Color.Black
                ishape.LineFormat.Width = 5

                ' Write the PPTX file to disk
                presentation.Save(dataDir & "EllipseShp2_out.pptx", SaveFormat.Pptx)
            End Using
			'ExEnd:FormattedEllipse	
        End Sub
    End Class
End Namespace