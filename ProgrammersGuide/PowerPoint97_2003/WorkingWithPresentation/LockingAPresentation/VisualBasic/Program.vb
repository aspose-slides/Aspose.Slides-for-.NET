'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Slides. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////
Imports System.IO

Imports Aspose.Slides

Namespace LockingAPresentation
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate presentation class
			Dim pres As New Presentation(dataDir & "demo.ppt")

			'Loop through all the slides in the presentation
			For Each sld As Slide In pres.Slides
				'Loop through all the shapes in the slide
				For Each shp As Shape In sld.Shapes
					'Lock each shape to be protected against the select
					shp.Protection = ShapeProtection.LockSelect
				Next shp
			Next sld

			'Write the presentation to disk
			pres.Write(dataDir & "demoLock.ppt")
		End Sub
	End Class
End Namespace