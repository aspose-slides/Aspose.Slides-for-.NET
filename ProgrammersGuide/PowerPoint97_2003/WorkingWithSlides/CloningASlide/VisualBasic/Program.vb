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

Namespace CloningASlide
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'********************************* Clone a Slide in Same Presentation ***********************************

			'Instantiate a Presentation object that represents a PPT file
			Dim pres1 As New Presentation(dataDir & "demo.ppt")


			'Accessing a slide using its slide position
			Dim slide As Slide = pres1.GetSlideByPosition(1)


			'Cloning the selected slide at the end of the same presentation file
			pres1.CloneSlide(slide, pres1.Slides.LastSlidePosition + 1)


			'Writing the presentation as a PPT file
			pres1.Write(dataDir & "CloneSlide1.ppt")

			'Instantiate a Presentation where the cloned slide will be added
			Dim pres2 As New Presentation(dataDir & "demo2.ppt")


			'Creating SortedList object that is used to store the temporary information
			'about the masters of PPT file. No value should be added to it.
			Dim sList As New System.Collections.SortedList()


			'Cloning the selected slide at the end of another presentation file
			pres1.CloneSlide(slide, pres2.Slides.LastSlidePosition + 1, pres2, sList)


			'Writing the presentation as a PPT file
			pres2.Write(dataDir & "CloneSlide2.ppt")



		End Sub
	End Class
End Namespace