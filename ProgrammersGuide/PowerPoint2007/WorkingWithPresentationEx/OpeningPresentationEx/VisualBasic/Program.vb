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

Namespace OpeningPresentationEx
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Opening the presentation file by passing the file path to the constructor of Presentation class
			Dim pres As New PresentationEx(dataDir & "demo.pptx")

			'Printing the total number of slides present in the presentation
			System.Console.WriteLine(pres.Slides.Count.ToString())

		End Sub
	End Class
End Namespace