'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Slides. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System.IO
Imports System.Drawing

Imports Aspose.Slides

Namespace SettingBackgroundColor
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'*************************** Setting the background color of Master Slide **************************

			'Instantiate a Presentation object that represents a PPT file
			Dim pres As New Presentation(dataDir & "demo.ppt")

			'Setting the background color to blue
			pres.Masters(0).Background.FillFormat.ForeColor = System.Drawing.Color.Blue

			'Writing the presentation as a PPT file
			pres.Write(dataDir & "MasterSlide.ppt")

			'*************************** Setting the background color of Normal Slide **************************

			'Instantiate a Presentation object that represents a PPT file
			pres = New Presentation(dataDir & "demo.ppt")

			'Accessing a slide using its slide position
			Dim slide As Slide = pres.GetSlideByPosition(1)


			'Disable following master background settings
			slide.FollowMasterBackground = False


			'Setting the background color to blue
			slide.Background.FillFormat.ForeColor = System.Drawing.Color.Blue


			'Writing the presentation as a PPT file
			pres.Write(dataDir & "NormalSlide.ppt")
		End Sub
	End Class
End Namespace