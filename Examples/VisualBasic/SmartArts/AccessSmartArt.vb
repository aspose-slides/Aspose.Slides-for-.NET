Imports Microsoft.VisualBasic
Imports System.IO
Imports Aspose.Slides
Imports System

Namespace Aspose.Slides.Examples.VisualBasic.SmartArts
    Public Class AccessSmartArt
        Public Shared Sub Run()
            ' ExStart:AccessSmartArt

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_SmartArts()

            ' Load the desired the presentation
            ' Load the desired the presentation
            Dim pres As New Presentation(dataDir & "AccessSmartArt.pptx")

            'Traverse through every shape inside first slide
            For Each shape As IShape In pres.Slides(0).Shapes

                ' Check if shape is of SmartArt type
                If TypeOf shape Is Aspose.Slides.SmartArt.SmartArt Then

                    'Typecast shape to SmartArt
                    Dim smart As Aspose.Slides.SmartArt.SmartArt = CType(shape, Aspose.Slides.SmartArt.SmartArt)

                    'Traverse through all nodes inside SmartArt
                    For i As Integer = 0 To smart.AllNodes.Count - 1
                        ' Accessing SmartArt node at index i
                        Dim node As Aspose.Slides.SmartArt.SmartArtNode = CType(smart.AllNodes(i), Aspose.Slides.SmartArt.SmartArtNode)

                        ' Printing the SmartArt node parameters
                        Dim outString As String = String.Format("i = {0}, Text = {1},  Level = {2}, Position = {3}", i, node.TextFrame.Text, node.Level, node.Position)
                        Console.WriteLine(outString)
                    Next i
                End If
            Next shape
            ' ExEnd:AccessSmartArt
        End Sub
    End Class
End Namespace