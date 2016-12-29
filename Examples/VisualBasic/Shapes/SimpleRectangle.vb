Imports Aspose.Slides.Export

Namespace Aspose.Slides.Examples.VisualBasic.Shapes
    Public Class SimpleRectangle
        Public Shared Sub Run()
			'ExStart:SimpleRectangle
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

                ' Add autoshape of rectangle type
                sld.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 150, 150, 50)

                'Write the PPTX file to disk
                pres.Save(dataDir & "RectShp1_out.pptx", SaveFormat.Pptx)
            End Using
			'ExEnd:SimpleRectangle
        End Sub
    End Class
End Namespace