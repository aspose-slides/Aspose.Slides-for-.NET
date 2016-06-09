'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Slides. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides

Namespace VisualBasic.Presentations
    Public Class ModifyBuiltinProperties
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            'Instantiate the Presentation class that represents the Presentation
            Dim pres As New Presentation(dataDir & "ModifyBuiltinProperties.pptx")

            'Create a reference to IDocumentProperties object associated with Presentation
            Dim dp As IDocumentProperties = pres.DocumentProperties

            'Set the builtin properties
            dp.Author = "Aspose.Slides for .NET"
            dp.Title = "Modifying Presentation Properties"
            dp.Subject = "Aspose Subject"
            dp.Comments = "Aspose Description"
            dp.Manager = "Aspose Manager"

            'Save your presentation to a file
            pres.Save(dataDir & "Updated_Document_Properties.pptx", Aspose.Slides.Export.SaveFormat.Pptx)

        End Sub
    End Class
End Namespace