Imports Microsoft.VisualBasic
Imports System.IO
Imports Aspose.Slides
Imports System

Namespace Aspose.Slides.Examples.VisualBasic.SmartArts
    Public Class AssistantNode
        Public Shared Sub Run()
            ' ExStart:AssistantNode
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_SmartArts()

            ' Creating a presentation instance
            Using pres As New Presentation(dataDir & "AssistantNode.pptx")
                'Traverse through every shape inside first slide
                For Each shape As IShape In pres.Slides(0).Shapes
                    ' Check if shape is of SmartArt type
                    If TypeOf shape Is Aspose.Slides.SmartArt.ISmartArt Then
                        'Typecast shape to SmartArtEx
                        Dim smart As Aspose.Slides.SmartArt.ISmartArt = CType(shape, Aspose.Slides.SmartArt.SmartArt)
                        'Traversing through all nodes of SmartArt shape

                        For Each node As Aspose.Slides.SmartArt.ISmartArtNode In smart.AllNodes
                            Dim tc As String = node.TextFrame.Text
                            ' Check if node is Assitant node
                            If node.IsAssistant Then
                                ' Setting Assitant node to false and making it normal node
                                node.IsAssistant = False
                            End If
                        Next node
                    End If
                Next shape
                ' Save Presentation
                pres.Save(dataDir & "ChangeAssitantNode_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx)
            End Using
            ' ExEnd:AssistantNode
        End Sub
    End Class
End Namespace