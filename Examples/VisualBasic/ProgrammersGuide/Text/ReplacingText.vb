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

Namespace VisualBasic.Text
    Public Class ReplacingText
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Text()

            'Instantiate Presentation class that represents PPTX//Instantiate Presentation class that represents PPTX
            Using pres As New Presentation(dataDir & "ReplacingText.pptx")

                'Access first slide
                Dim sld As ISlide = pres.Slides(0)

                'Iterate through shapes to find the placeholder
                For Each shp As IShape In sld.Shapes
                    If shp.Placeholder IsNot Nothing Then
                        'Change the text of each placeholder
                        CType(shp, IAutoShape).TextFrame.Text = "This is Placeholder"
                    End If
                Next shp

                'Save the PPTX to Disk
                pres.Save(dataDir & "output.pptx", Aspose.Slides.Export.SaveFormat.Pptx)
            End Using

        End Sub
    End Class
End Namespace