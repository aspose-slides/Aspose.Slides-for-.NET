'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Slides. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////
Imports System.IO

Imports Aspose.Slides

Namespace OpenSimple
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Opening the presentation file by passing the file path to the constructor
			'of the Presentation class
			Dim pres As New Presentation(dataDir & "simple.ppt")

			'Printing the total number of slides in the presentation
			System.Console.WriteLine("Number of slides in simple presentation are : " & pres.Slides.Count.ToString())
		End Sub
	End Class
End Namespace