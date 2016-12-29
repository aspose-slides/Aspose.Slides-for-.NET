Imports Aspose.Slides.Export

Namespace Aspose.Slides.Examples.VisualBasic.Shapes
    Public Class RotatingShapes
        Public Shared Sub Run()
			'ExStart:RotatingShapes	
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not IsExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If

            ' Instantiate PrseetationEx class that represents the PPTX
            Using pres As New Presentation()

                ' Get the first slide
                Dim sld As ISlide = pres.Slides(0)

                ' Add autoshape of rectangle type
                Dim shp As IShape = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 150, 75, 150)

                ' Rotate the shape to 90 degree
                shp.Rotation = 90

                'Write the PPTX file to disk
                pres.Save(dataDir & "RectShpRot_out.pptx", SaveFormat.Pptx)
            End Using
			'ExEnd:RotatingShapes	
        End Sub
    End Class
End Namespace