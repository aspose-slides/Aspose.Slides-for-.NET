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
    Public Class RemoveNode
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_SmartArts()

            'Load the desired the presentation
            Using pres As New Presentation(dataDir & "RemoveNode.pptx")

                'Traverse through every shape inside first slide
                For Each shape As IShape In pres.Slides(0).Shapes

                    'Check if shape is of SmartArt type
                    If TypeOf shape Is ISmartArt Then
                        'Typecast shape to SmartArtEx
                        Dim smart As ISmartArt = CType(shape, ISmartArt)

                        If smart.AllNodes.Count > 0 Then
                            'Accessing SmartArt node at index 0
                            Dim node As ISmartArtNode = smart.AllNodes(0)

                            'Removing the selected node
                            smart.AllNodes.RemoveNode(node)

                        End If
                    End If
                Next shape

                'Save Presentation
                pres.Save(dataDir & "RemoveSmartArtNode.pptx", Aspose.Slides.Export.SaveFormat.Pptx)
            End Using


        End Sub
    End Class
End Namespace