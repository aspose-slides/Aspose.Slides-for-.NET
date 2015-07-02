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

Namespace VisualBasic.Presentations
    Public Class OpenPasswordPresentation
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            'creating instance of load options to set the presentation access password
            Dim loadOptions As New Aspose.Slides.LoadOptions()

            'Setting the access password
            loadOptions.Password = "pass"

            'Opening the presentation file by passing the file path and load options to the constructor of Presentation class
            Dim pres As New Presentation(dataDir & "OpenPasswordPresentation.pptx", loadOptions)

            'Printing the total number of slides present in the presentation
            System.Console.WriteLine(pres.Slides.Count.ToString())
        End Sub
    End Class
End Namespace