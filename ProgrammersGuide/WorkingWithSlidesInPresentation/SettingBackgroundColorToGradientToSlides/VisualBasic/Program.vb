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

Namespace SettingBackgroundColorToGradientToSlides
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate the Presentation class that represents the presentation file
			Using pres As New Presentation(dataDir & "Aspose.pptx")

				'Apply Gradiant effect to the Background
				pres.Slides(0).Background.Type = BackgroundType.OwnBackground
				pres.Slides(0).Background.FillFormat.FillType = FillType.Gradient
				pres.Slides(0).Background.FillFormat.GradientFormat.TileFlip = TileFlip.FlipBoth

				'Write the presentation to disk
				pres.Save(dataDir & "ContentBG_Grad.pptx", SaveFormat.Pptx)

			End Using

		End Sub
	End Class
End Namespace