Imports System

Namespace Aspose.Slides.Examples.VisualBasic.Shapes
    Public Class ConnectorLineAngle
        Public Shared Sub Run()
			'ExStart:ConnectorLineAngle
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            Dim pres As New Presentation(dataDir & "ConnectorLineAngle.pptx")
            Dim slide As Slide = CType(pres.Slides(0), Slide)
            Dim shape As Shape
            For i As Integer = 0 To slide.Shapes.Count - 1
                Dim dir As Double = 0.0
                shape = CType(slide.Shapes(i), Shape)
                If TypeOf shape Is AutoShape Then
                    Dim ashp As AutoShape = CType(shape, AutoShape)
                    If ashp.ShapeType = ShapeType.Line Then
                        dir = getDirection(ashp.Width, ashp.Height, Convert.ToBoolean(ashp.Frame.FlipH), Convert.ToBoolean(ashp.Frame.FlipV))
                    End If
                ElseIf TypeOf shape Is Connector Then
                    Dim ashp As Connector = CType(shape, Connector)
                    dir = getDirection(ashp.Width, ashp.Height, Convert.ToBoolean(ashp.Frame.FlipH), Convert.ToBoolean(ashp.Frame.FlipV))
                End If

                Console.WriteLine(dir)
            Next i

        End Sub
        Public Shared Function getDirection(ByVal w As Single, ByVal h As Single, ByVal flipH As Boolean, ByVal flipV As Boolean) As Double
            Dim endLineX As Single = w * (If(flipH, -1, 1))
            Dim endLineY As Single = h * (If(flipV, -1, 1))
            Dim endYAxisX As Single = 0
            Dim endYAxisY As Single = h
            Dim angle As Double = (Math.Atan2(endYAxisY, endYAxisX) - Math.Atan2(endLineY, endLineX))
            If angle < 0 Then
                angle += 2 * Math.PI
            End If
            Return angle * 180.0 / Math.PI
        End Function
		
		'ExEnd:ConnectorLineAngle
    End Class

End Namespace