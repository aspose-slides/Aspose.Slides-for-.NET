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

Namespace AddingSlidesToPresentationEx
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate PresentationEx class that represents the PPTX file
			Dim pres As New PresentationEx()

			'Instantiate SlideExCollection calss
			Dim slds As SlideExCollection = pres.Slides

			'Add an empty slide to the SlidesEx collection
			slds.AddEmptySlide(pres.LayoutSlides(0))

			'Do some work on the newly added slide

			'Save the PPTX file to the Disk
			pres.Write(dataDir & "EmptySlide.pptx")


		End Sub
	End Class
End Namespace