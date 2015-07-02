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
Imports Aspose.Slides.SmartArt

Namespace VisualBasic.SmartArts
    Public Class CreateSmartArtShape
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_SmartArts()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not IsExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If
            'Instantiate the presentation
            Using pres As New Presentation()

                'Access the presentation slide
                Dim slide As ISlide = pres.Slides(0)

                'Add Smart Art Shape
                Dim smart As ISmartArt = slide.Shapes.AddSmartArt(0, 0, 400, 400, SmartArtLayoutType.BasicBlockList)

                'Saving presentation
                pres.Save(dataDir & "SimpleSmartArt.pptx", Aspose.Slides.Export.SaveFormat.Pptx)
            End Using
        End Sub
    End Class
End Namespace