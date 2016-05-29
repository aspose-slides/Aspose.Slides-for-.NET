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
Imports Aspose.Slides.SmartArt

Namespace VisualBasic.SmartArts
    Public Class RemoveNodeSpecificPosition
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_SmartArts()

            'Load the desired the presentation
            Dim pres As New Presentation(dataDir & "RemoveNodesSpecificPosition.pptx")

            'Traverse through every shape inside first slide
            For Each shape As IShape In pres.Slides(0).Shapes

                'Check if shape is of SmartArt type
                If TypeOf shape Is SmartArt Then
                    'Typecast shape to SmartArt
                    Dim smart As SmartArt = CType(shape, SmartArt)

                    If smart.AllNodes.Count > 0 Then
                        'Accessing SmartArt node at index 0
                        Dim node As ISmartArtNode = smart.AllNodes(0)

                        If node.ChildNodes.Count >= 2 Then
                            'Removing the child node at position 1
                            CType(node.ChildNodes, SmartArtNodeCollection).RemoveNode(1)
                        End If

                    End If
                End If

            Next shape

            'Save Presentation
            pres.Save(dataDir & "RemoveSmartArtNodeByPosition.pptx", Aspose.Slides.Export.SaveFormat.Pptx)

        End Sub
    End Class
End Namespace