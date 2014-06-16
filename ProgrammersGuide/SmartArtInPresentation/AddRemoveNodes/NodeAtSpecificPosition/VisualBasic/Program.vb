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

Namespace NodeAtSpecificPosition
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Create directory if it is not already present.
			Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
			If (Not IsExists) Then
				System.IO.Directory.CreateDirectory(dataDir)
			End If

			'Creating a presentation instance
			Dim pres As New Presentation()

			'Access the presentation slide
			Dim slide As ISlide = pres.Slides(0)

			'Add Smart Art IShape
			Dim smart As ISmartArt = slide.Shapes.AddSmartArt(0, 0, 400, 400, SmartArtLayoutType.StackedList)

			'Accessing the SmartArt node at index 0
			Dim node As ISmartArtNode = smart.AllNodes(0)

			'Adding new child node at position 2 in parent node
			Dim chNode As SmartArtNode = CType((CType(node.ChildNodes, SmartArtNodeCollection)).AddNodeByPosition(2), SmartArtNode)

			'Add Text
			chNode.TextFrame.Text = "Sample Text Added"

			'Save Presentation
			pres.Save(dataDir & "AddSmartArtNodeByPosition.pptx", Aspose.Slides.Export.SaveFormat.Pptx)


		End Sub
	End Class
End Namespace