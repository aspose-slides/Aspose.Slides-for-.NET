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

Namespace UsingSlideIndex
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate a Presentation object that represents a presentation file
			Using pres As New Presentation(dataDir & "Aspose.pptx")

				'Removing a slide using its slide index
				pres.Slides.RemoveAt(0)


				'Writing the presentation file
				pres.Write(dataDir & "modified.pptx")

			End Using
		End Sub
	End Class
End Namespace