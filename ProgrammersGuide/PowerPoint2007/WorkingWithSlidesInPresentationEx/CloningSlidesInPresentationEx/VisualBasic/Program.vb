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
Imports Aspose.Slides.Pptx

Namespace CloningSlidesInPresentationEx
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate PresentationEx class that represents a PPTX file
			Dim pres As New PresentationEx(dataDir & "demo.pptx")

			'Clone the desired slide to the end of the collection of slides in the same PPTX
			Dim slds As SlideExCollection = pres.Slides
			slds.AddClone(pres.Slides(0))

			'Write the modified pptx to disk
			pres.Write(dataDir & "demo_cloned.pptx")

		End Sub
	End Class
End Namespace