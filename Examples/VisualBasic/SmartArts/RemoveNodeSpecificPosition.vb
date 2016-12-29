Imports Microsoft.VisualBasic
Imports System.IO
Imports Aspose.Slides

Namespace Aspose.Slides.Examples.VisualBasic.SmartArts
    Public Class RemoveNodeSpecificPosition
        Public Shared Sub Run()
            ' ExStart:RemoveNodeSpecificPosition
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_SmartArts()

            ' Load the desired the presentation  
            Dim pres As New Presentation(dataDir & "RemoveNodeSpecificPosition.pptx")

            'Traverse through every shape inside first slide
            For Each shape As IShape In pres.Slides(0).Shapes

                ' Check if shape is of SmartArt type
                If TypeOf shape Is Aspose.Slides.SmartArt.SmartArt Then
                    'Typecast shape to SmartArt
                    Dim smart As Aspose.Slides.SmartArt.SmartArt = CType(shape, Aspose.Slides.SmartArt.SmartArt)

                    If smart.AllNodes.Count > 0 Then
                        ' Accessing SmartArt node at index 0
                        Dim node As Aspose.Slides.SmartArt.ISmartArtNode = smart.AllNodes(0)

                        If node.ChildNodes.Count >= 2 Then
                            ' Removing the child node at position 1
                            CType(node.ChildNodes, Aspose.Slides.SmartArt.SmartArtNodeCollection).RemoveNode(1)
                        End If

                    End If
                End If
            Next shape
            ' Save Presentation
            pres.Save(dataDir & "RemoveSmartArtNodeByPosition_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx)
            ' ExEnd:RemoveNodeSpecificPosition
        End Sub
    End Class
End Namespace