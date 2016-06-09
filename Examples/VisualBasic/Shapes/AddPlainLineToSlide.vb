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
Imports Aspose.Slides.Export

Namespace VisualBasic.Shapes
    Public Class AddPlainLineToSlide
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not IsExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If

            'Instantiate PresentationEx class that represents the PPTX file
            Using pres As New Presentation()
                'Get the first slide
                Dim sld As ISlide = pres.Slides(0)

                'Add an autoshape of type line
                sld.Shapes.AddAutoShape(ShapeType.Line, 50, 150, 300, 0)

                'Write the PPTX to Disk
                pres.Save(dataDir & "LineShape1.pptx", SaveFormat.Pptx)
            End Using
        End Sub
    End Class
End Namespace