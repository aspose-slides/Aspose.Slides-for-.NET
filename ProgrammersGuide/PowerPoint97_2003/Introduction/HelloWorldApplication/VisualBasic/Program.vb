'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Slides. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////
Imports System.IO

Imports Aspose.Slides

Namespace HelloWorldApplication
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate a Presentation object that represents a PPT file
			Dim pres As New Presentation()

			'Create a License object
			Dim license As New License()

			'Set the license of Aspose.Slides for .NET to avoid the evaluation limitations
			'license.SetLicense("Aspose.Slides.lic");

			'Adding an empty slide to the presentation and getting the reference of
			'that empty slide
			Dim slide As Slide = pres.AddEmptySlide()

			'Adding a rectangle (X=2400, Y=1800, Width=1000 & Height=500) to the slide
			Dim rect As Aspose.Slides.Rectangle = slide.Shapes.AddRectangle(2400, 1800, 1000, 500)

			'Hiding the lines of rectangle
			rect.LineFormat.ShowLines = False

			'Adding a text frame to the rectangle with "Hello World" as a default text
			rect.AddTextFrame("Hello World")

			'Removing the first slide of the presentation which is always added by
			'Aspose.Slides for .NET by default while creating the presentation
			pres.Slides.RemoveAt(0)

			'Writing the presentation as a PPT file
			pres.Write(dataDir & "hello.ppt")

			' Inform user about creation of hello.ppt
			System.Console.WriteLine("hello.ppt created successfully!")
		End Sub
	End Class
End Namespace