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

Namespace RemovingWriteProtection
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Opening the presentation file
			Dim pres As New Presentation(dataDir & "demoWriteProtected.pptx")


			'Checking if presentation is write protected
			If pres.ProtectionManager.IsWriteProtected Then
				'Removing Write protection
				pres.ProtectionManager.RemoveWriteProtection()
			End If

			'Saving presentation
			pres.Write(dataDir & "newDemo.pptx")
		End Sub
	End Class
End Namespace