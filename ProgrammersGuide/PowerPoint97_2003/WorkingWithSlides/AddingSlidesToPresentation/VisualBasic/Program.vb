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

Namespace AddingSlidesToPresentation
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'********************** Adding Empty Slide in a Presentation *****************************

			'Instantiate a Presentation object that represents a PPT file
			Dim pres As New Presentation(dataDir & "demo.ppt")

			'Adding an empty slide to the presentation and getting the reference of
			'that empty slide
			Dim slide1 As Slide = pres.AddEmptySlide()

			'Writing the presentation as a PPT file
			pres.Write(dataDir & "EmptySlide.ppt")

			'********************** Adding Body Slide in a Presentation *****************************

			'Instantiate a Presentation object that represents a PPT file
			pres = New Presentation(dataDir & "demo.ppt")

			'Adding a body slide to the presentation and getting the reference of
			'that body slide
			Dim slide2 As Slide = pres.AddBodySlide()

			'Writing the presentation as a PPT file
			pres.Write(dataDir & "BodySlide.ppt")

			'********************** Adding Double Body Slide in a Presentation *****************************

			'Instantiate a Presentation object that represents a PPT file
			pres = New Presentation(dataDir & "demo.ppt")

			'Adding a double body slide to the presentation and getting the reference of
			'that double body slide
			Dim slide3 As Slide = pres.AddDoubleBodySlide()

			'Writing the presentation as a PPT file
			pres.Write(dataDir & "DoubleBodySlide.ppt")

			'********************** Adding Header Slide in a Presentation *****************************

			'Instantiate a Presentation object that represents a PPT file
			pres = New Presentation(dataDir & "demo.ppt")

			'Adding a header slide to the presentation and getting the reference of
			'that header slide
			Dim slide4 As Slide = pres.AddHeaderSlide()

			'Writing the presentation as a PPT file
			pres.Write(dataDir & "HeaderSlide.ppt")

			'********************** Adding Title Slide in a Presentation *****************************

			'Instantiate a Presentation object that represents a PPT file
			pres = New Presentation(dataDir & "demo.ppt")

			'Adding a title slide to the presentation and getting the reference of
			'that title slide
			Dim slide5 As Slide = pres.AddTitleSlide()

			'Writing the presentation as a PPT file
			pres.Write(dataDir & "TitleSlide.ppt")


		End Sub
	End Class
End Namespace