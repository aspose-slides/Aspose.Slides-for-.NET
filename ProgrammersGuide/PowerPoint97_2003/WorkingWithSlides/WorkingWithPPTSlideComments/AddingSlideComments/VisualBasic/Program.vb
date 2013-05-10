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
Imports System

Namespace AddingSlideComments
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			Dim pres As New Presentation()

			'Getting first slide
			Dim slide As Slide = pres.Slides(0)

			'Adding Autthor
			Dim author As CommentAuthor = pres.CommentAuthors.AddAuthor("Aspose")


			'Position of comments
			Dim point As New Point()
			point.X = 100
			point.Y = 100

			'Adding Slide comments
			slide.SlideComments.AddComment(author, "AP", "Hello Aspose, this is slide comment", DateTime.Now, point)

			'Adding Empty slide
			slide = pres.AddEmptySlide()

			'Position of comments
			Dim point2 As New Point()
			point2.X = 500
			point2.Y = 1400

			'Adding Slide comments
			slide.SlideComments.AddComment(author, "AP", "Hello Aspose, this is second slide comment", DateTime.Now, point2)

			Dim comments As CommentCollection = slide.SlideComments
			'Accessin the comment at index 0 for slide 1
			Dim str As String = comments(0).Text

			pres.Write(dataDir & "Comments.ppt")
		End Sub
	End Class
End Namespace