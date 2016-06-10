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
Imports System

Namespace VisualBasic.SmartArts
    Public Class AssistantNode
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_SmartArts()

            'Creating a presentation instance
            Using pres As New Presentation(dataDir & "AssistantNode.pptx")
                'Traverse through every shape inside first slide
                For Each shape As IShape In pres.Slides(0).Shapes
                    'Check if shape is of SmartArt type
                    If TypeOf shape Is ISmartArt Then
                        'Typecast shape to SmartArtEx
                        Dim smart As ISmartArt = CType(shape, SmartArt)
                        'Traversing through all nodes of SmartArt shape

                        For Each node As ISmartArtNode In smart.AllNodes
                            Dim tc As String = node.TextFrame.Text
                            'Check if node is Assitant node
                            If node.IsAssistant Then
                                'Setting Assitant node to false and making it normal node
                                node.IsAssistant = False
                            End If
                        Next node
                    End If
                Next shape
                'Save Presentation
                pres.Save(dataDir & "ChangeAssitantNode.pptx", Aspose.Slides.Export.SaveFormat.Pptx)
            End Using
            'Creating a presentation instance


        End Sub
    End Class
End Namespace