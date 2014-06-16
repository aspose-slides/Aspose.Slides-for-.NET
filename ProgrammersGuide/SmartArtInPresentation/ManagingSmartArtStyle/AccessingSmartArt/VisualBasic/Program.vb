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
Imports System

Namespace AccessingSmartArt
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Load the desired the presentation
			Using pres As New Presentation(dataDir & "SimpleSmartArt.pptx")

				'Traverse through every shape inside first slide
				For Each shape As IShape In pres.Slides(0).Shapes
					'Check if shape is of SmartArt type
					If TypeOf shape Is ISmartArt Then
						'Typecast shape to SmartArtEx
						Dim smart As ISmartArt = CType(shape, ISmartArt)

						'Checking SmartArt Layout
						If smart.Layout = SmartArtLayoutType.BasicBlockList Then
							Console.WriteLine("Layout Type:" & smart.Layout)
						End If

					End If

				Next shape
			End Using

		End Sub
	End Class
End Namespace