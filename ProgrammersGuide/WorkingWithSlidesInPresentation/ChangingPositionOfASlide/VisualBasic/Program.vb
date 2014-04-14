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
Imports Aspose.Slides.Export

Namespace ChangingPositionOfASlide
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate Presentation class to load the source presentation file
			Using pres As New Presentation(dataDir & "Aspose.pptx")
				'Get the slide whose position is to be changed
				Dim sld As ISlide = pres.Slides(0)

				'Set the new position for the slide
				sld.SlideNumber = 2

				'Write the presentation to disk
				pres.Save(dataDir & "Aspose_out.pptx", SaveFormat.Pptx)

			End Using
		End Sub
	End Class
End Namespace