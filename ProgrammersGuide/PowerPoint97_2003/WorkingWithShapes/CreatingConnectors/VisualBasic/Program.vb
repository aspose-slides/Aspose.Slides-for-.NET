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

Namespace CreatingConnectors
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Create directory if it is not already present.
			Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
			If (Not IsExists) Then
				System.IO.Directory.CreateDirectory(dataDir)
			End If

			'Instantiate a Presentation object with new empty PPT file
			Dim pres As New Presentation()


			'Accessing a slide using its slide position
			Dim slide As Slide = pres.GetSlideByPosition(1)


			'Creating 4 rectangles
			Dim root As Aspose.Slides.Rectangle = CreateRectangle(slide, 500, 500, 2760, 500, "Connectors")
			Dim straight As Aspose.Slides.Rectangle = CreateRectangle(slide, 200, 3500, 2000, 400, "Straight")
			Dim elbow As Aspose.Slides.Rectangle = CreateRectangle(slide, 3500, 1500, 2000, 400, "Elbow")
			Dim curve As Aspose.Slides.Rectangle = CreateRectangle(slide, 3000, 2500, 2000, 400, "Curve")

			'Create straight connector
			CreateConnector(slide, ConnectorType.Straight, root, 2, straight, 0)


			'Create elbow connector
			CreateConnector(slide, ConnectorType.Elbow, root, 3, elbow, 0)


			'Create curve connector
			CreateConnector(slide, ConnectorType.Curve, root, 2, curve, 1)

			pres.Write(dataDir & "output.ppt")

		End Sub

		Private Shared Function CreateRectangle(ByVal slide As Slide, ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer, ByVal text As String) As Aspose.Slides.Rectangle
			' Create new Rectangle shape on a slide
			Dim rectangle As Aspose.Slides.Rectangle = slide.Shapes.AddRectangle(x, y, w, h)


			' Set format of lines for the rectangle
			rectangle.LineFormat.Width = 5
			rectangle.LineFormat.ForeColor = Color.Red


			' Add centered text
			rectangle.AddTextFrame(text)
			Dim tf As TextFrame = rectangle.TextFrame
			tf.Paragraphs(0).Alignment = TextAlignment.Center
			tf.Paragraphs(0).Portions(0).FontBold = True
			tf.Paragraphs(0).Portions(0).FontHeight = 36


			' Return created shape
			Return rectangle
		End Function


		Private Shared Function CreateConnector(ByVal slide As Slide, ByVal type As ConnectorType, ByVal shape1 As Shape, ByVal connPoint1 As Integer, ByVal shape2 As Shape, ByVal connPoint2 As Integer) As Connector
			' Add new connector with some random default coordinates
			Dim connector As Connector = slide.Shapes.AddConnector(type, New Point(500, 500), New Point(1000, 1000))


			' Connect connector with 2 shapes
			connector.ConnectBegin(shape1, connPoint1)
			connector.ConnectEnd(shape2, connPoint2)


			' Set connector style
			connector.LineFormat.ForeColor = Color.Blue
			connector.LineFormat.Width = 5
			connector.LineFormat.EndArrowheadStyle = LineArrowheadStyle.Open


			' Return created connector
			Return connector
		End Function


	End Class
End Namespace