'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Slides. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports System.Drawing
Imports Microsoft.VisualBasic
Imports System.IO
Imports System.Runtime.InteropServices.ComTypes

Imports Aspose.Slides
Imports Aspose.Slides.Export

Namespace VisualBasic.Presentations
    Public Class AccessBuiltinProperties
        Public Shared Sub Run()

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            'Instantiate the Presentation class that represents the presentation
            Dim pres As New Presentation(dataDir & "AccessBuiltin Properties.pptx")

            'Create a reference to IDocumentProperties object associated with Presentation
            Dim dp As IDocumentProperties = pres.DocumentProperties

            'Display the builtin properties
            System.Console.WriteLine("Category : " & dp.Category)
            System.Console.WriteLine("Current Status : " & dp.ContentStatus)
            System.Console.WriteLine("Creation Date : " & dp.CreatedTime)
            System.Console.WriteLine("Author : " & dp.Author)
            System.Console.WriteLine("Description : " & dp.Comments)
            System.Console.WriteLine("KeyWords : " & dp.Keywords)
            System.Console.WriteLine("Last Modified By : " & dp.LastSavedBy)
            System.Console.WriteLine("Supervisor : " & dp.Manager)
            System.Console.WriteLine("Modified Date : " & dp.LastSavedTime)
            System.Console.WriteLine("Presentation Format : " & dp.PresentationFormat)
            System.Console.WriteLine("Last Print Date : " & dp.LastPrinted)
            System.Console.WriteLine("Is Shared between producers : " & dp.SharedDoc)
            System.Console.WriteLine("Subject : " & dp.Subject)
            System.Console.WriteLine("Title : " & dp.Title)

        End Sub
    End Class
End Namespace