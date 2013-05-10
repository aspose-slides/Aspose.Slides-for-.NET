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

Namespace SettingBackgroundToImage
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate a Presentation object that represents a PPT file
			Dim pres As New Presentation(dataDir & "demo.ppt")


			'Accessing a slide using its slide position
			Dim slide As Slide = pres.GetSlideByPosition(1)


			'Disable following master background settings
			slide.FollowMasterBackground = False


			'Setting the fill type of the background to picture
			slide.Background.FillFormat.Type = FillType.Picture


			'Creating a picture object that will be used as a slide background
			Dim pic As New Aspose.Slides.Picture(pres, dataDir & "logo.jpg")


			'Adding the picture object to pictures collection of the presentation
			'After the picture object is added, the picture is given a unique picture Id
			Dim picId As Integer = pres.Pictures.Add(pic)


			'Setting the picture Id of the slide background to the Id of the picture object
			slide.Background.PictureId = picId


			'Writing the presentation as a PPT file
			pres.Write(dataDir & "modified.ppt")
		End Sub
	End Class
End Namespace