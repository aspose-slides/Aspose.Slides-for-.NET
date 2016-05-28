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
    Public Class Convert_XPS_Options
        Public Shared Sub Run()

            ' For complete examples and data files, please go to https://github.com/aspose-slides/Aspose.Slides-for-.NET

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            ' Instantiate a Presentation object that represents a presentation file
            Using presentation As New Presentation(dataDir & "Convert_XPS_Options.pptx")

                ' Instantiate the TiffOptions class
                Dim options As New Aspose.Slides.Export.XpsOptions()

                ' Save MetaFiles as PNG
                options.SaveMetafilesAsPng = True

                ' Save the presentation to XPS document
                presentation.Save(dataDir & "XPS_With_Options.xps", Aspose.Slides.Export.SaveFormat.Xps, options)

            End Using
        End Sub
    End Class
End Namespace