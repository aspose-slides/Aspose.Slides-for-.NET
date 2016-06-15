Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
 Imports AsposeSlides = Aspose.Slides.SmartArt

Namespace Aspose.Slides.Examples.VisualBasic.SmartArts
    Public Class RemoveNodeSpecificPosition
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_SmartArts()

            'Load the desired the presentation  
            Dim pres As New Presentation(dataDir & "RemoveNodeSpecificPosition.pptx")

            'Traverse through every shape inside first slide
            For Each shape As IShape In pres.Slides(0).Shapes

                'Check if shape is of SmartArt type
                If TypeOf shape Is AsposeSlides.SmartArt Then
                    'Typecast shape to SmartArt
                    Dim smart As AsposeSlides.SmartArt = CType(shape, AsposeSlides.SmartArt)

                    If smart.AllNodes.Count > 0 Then
                        'Accessing SmartArt node at index 0
                        Dim node As AsposeSlides.ISmartArtNode = smart.AllNodes(0)

                        If node.ChildNodes.Count >= 2 Then
                            'Removing the child node at position 1
                            CType(node.ChildNodes, AsposeSlides.SmartArtNodeCollection).RemoveNode(1)
                        End If

                    End If
                End If

            Next shape

            'Save Presentation
            pres.Save(dataDir & "RemoveSmartArtNodeByPosition.pptx", Aspose.Slides.Export.SaveFormat.Pptx)

        End Sub
    End Class
End Namespace