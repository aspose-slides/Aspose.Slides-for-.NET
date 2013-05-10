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
Imports System

Namespace AccessingSlideComments
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			Dim pres As New Presentation(dataDir & "Comments.ppt")


			Dim i As Integer = 1
			For Each slide As Slide In pres.Slides
				For Each comment As Comment In slide.SlideComments
					Console.WriteLine("Slide :" & i.ToString() & " has comment: " & comment.Text & " with Author: " & comment.Author.Name & " posted on time :" & comment.CreatedTime + Constants.vbLf)
					i += 1
				Next comment
			Next slide


		End Sub
	End Class
End Namespace