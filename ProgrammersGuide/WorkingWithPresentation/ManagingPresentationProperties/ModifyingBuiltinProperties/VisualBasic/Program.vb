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

Namespace ModifyingBuiltinProperties
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate the Presentation class that represents the Presentation
			Dim pres As New Presentation(dataDir & "Aspose.pptx")

			'Create a reference to IDocumentProperties object associated with Presentation
			Dim dp As IDocumentProperties = pres.DocumentProperties

			'Set the builtin properties
			dp.Author = "Aspose.Slides for .NET"
			dp.Title = "Modifying Presentation Properties"
			dp.Subject = "Aspose Subject"
			dp.Comments = "Aspose Description"
			dp.Manager = "Aspose Manager"

			'Save your presentation to a file
			pres.Write(dataDir & "DocProps.pptx")

		End Sub
	End Class
End Namespace