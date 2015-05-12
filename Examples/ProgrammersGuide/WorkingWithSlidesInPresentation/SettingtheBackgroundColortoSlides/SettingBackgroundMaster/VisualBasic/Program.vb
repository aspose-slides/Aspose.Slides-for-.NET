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
Imports Aspose.Slides.Export
Imports System.Drawing

Namespace SettingBackgroundMaster
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Create directory if it is not already present.
			Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
			If (Not IsExists) Then
				System.IO.Directory.CreateDirectory(dataDir)
			End If

			'Instantiate the Presentation class that represents the presentation file
			Using pres As New Presentation()

				'Set the background color of the Master ISlide to Forest Green
				pres.Masters(0).Background.Type = BackgroundType.OwnBackground
				pres.Masters(0).Background.FillFormat.FillType = FillType.Solid
				pres.Masters(0).Background.FillFormat.SolidFillColor.Color = Color.ForestGreen

				'Write the presentation to disk
				pres.Save(dataDir & "Aspose.pptx", SaveFormat.Pptx)

			End Using

		End Sub
	End Class
End Namespace