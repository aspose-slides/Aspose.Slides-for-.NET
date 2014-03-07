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

Namespace AccessingModifyingProperties
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instanciate the Presentation class that represents the PPTX
			Dim pres As New Presentation(dataDir & "Aspose.pptx")

			'Create a reference to DocumentProperties object associated with Prsentation
			Dim dp As IDocumentProperties = pres.DocumentProperties


			'Access and modify custom properties
			For i As Integer = 0 To dp.Count - 1
				'Display names and values of custom properties
				System.Console.WriteLine("Custom Property Name : " & dp.GetPropertyName(i))
				System.Console.WriteLine("Custom Property Value : " & dp(dp.GetPropertyName(i)))

				'Modify values of custom properties
				dp(dp.GetPropertyName(i)) = "New Value " & (i + 1)
			Next i

			'Save your presentation to a file
			pres.Write(dataDir & "CustomDemoModified.pptx")
		End Sub
	End Class
End Namespace