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
Imports Aspose.Slides.SmartArt

Namespace AddingNodes
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Load the desired the presentation//Load the desired the presentation
			Dim pres As New Presentation(dataDir & "SimpleSmartArt.pptx")

			'Traverse through every shape inside first slide
			For Each shape As IShape In pres.Slides(0).Shapes

				'Check if shape is of SmartArt type
				If TypeOf shape Is SmartArt Then

					'Typecast shape to SmartArt
					Dim smart As SmartArt = CType(shape, SmartArt)

					'Adding a new SmartArt Node
					Dim TemNode As SmartArtNode = CType(smart.AllNodes.AddNode(), SmartArtNode)

					'Adding text
					TemNode.TextFrame.Text = "Test"

					'Adding new child node in parent node. It  will be added in the end of collection
					Dim newNode As SmartArtNode = CType(TemNode.ChildNodes.AddNode(), SmartArtNode)

					'Adding text
					newNode.TextFrame.Text = "New Node Added"

				End If
			Next shape

			'Saving Presentation
			pres.Save(dataDir & "AddSmartArtNode.pptx", Aspose.Slides.Export.SaveFormat.Pptx)


		End Sub
	End Class
End Namespace