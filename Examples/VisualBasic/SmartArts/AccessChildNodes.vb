Imports System
Imports Aspose.Slides.SmartArt

Namespace Aspose.Slides.Examples.VisualBasic.SmartArts
    Public Class AccessChildNodes
        Public Shared Sub Run()
            'ExStart:AccessChildNodes
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_SmartArts()

            ' Load the desired the presentation
            Dim pres As New Presentation(dataDir & "AccessChildNodes.pptx")

            'Traverse through every shape inside first slide
            For Each shape As IShape In pres.Slides(0).Shapes

                ' Check if shape is of SmartArt type
                If TypeOf shape Is Aspose.Slides.SmartArt.SmartArt Then

                    'Typecast shape to SmartArt
                    Dim smart As Aspose.Slides.SmartArt.SmartArt = CType(shape, Aspose.Slides.SmartArt.SmartArt)

                    'Traverse through all nodes inside SmartArt
                    For i As Integer = 0 To smart.AllNodes.Count - 1
                        ' Accessing SmartArt node at index i
                        Dim node0 As SmartArtNode = CType(smart.AllNodes(i), SmartArtNode)

                        'Traversing through the child nodes in SmartArt node at index i
                        For j As Integer = 0 To node0.ChildNodes.Count - 1
                            ' Accessing the child node in SmartArt node
                            Dim node As SmartArtNode = CType(node0.ChildNodes(j), SmartArtNode)

                            ' Printing the SmartArt child node parameters
                            Dim outString As String = String.Format("j = {0}, Text = {1},  Level = {2}, Position = {3}", j, node.TextFrame.Text, node.Level, node.Position)
                            Console.WriteLine(outString)
                        Next j
                    Next i
                End If
            Next shape
            'ExEnd:AccessChildNodes
        End Sub
    End Class
End Namespace