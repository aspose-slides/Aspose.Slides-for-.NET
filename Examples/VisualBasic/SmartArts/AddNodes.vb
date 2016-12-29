Imports Microsoft.VisualBasic
Imports System.IO
Imports Aspose.Slides
 
Namespace Aspose.Slides.Examples.VisualBasic.SmartArts
    Public Class AddNodes
        Public Shared Sub Run()
            ' ExStart:AddNodes
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_SmartArts()

            ' Load the desired the presentation//Load the desired the presentation
            Dim pres As New Presentation(dataDir & "AddNodes.pptx")

            'Traverse through every shape inside first slide
            For Each shape As IShape In pres.Slides(0).Shapes

                ' Check if shape is of SmartArt type
                If TypeOf shape Is Aspose.Slides.SmartArt.SmartArt Then

                    'Typecast shape to SmartArt
                    Dim smart As Aspose.Slides.SmartArt.SmartArt = CType(shape, Aspose.Slides.SmartArt.SmartArt)

                    ' Adding a new SmartArt Node
                    Dim TemNode As Aspose.Slides.SmartArt.SmartArtNode = CType(smart.AllNodes.AddNode(), Aspose.Slides.SmartArt.SmartArtNode)

                    ' Adding text
                    TemNode.TextFrame.Text = "Test"

                    ' Adding new child node in parent node. It  will be added in the end of collection
                    Dim newNode As Aspose.Slides.SmartArt.SmartArtNode = CType(TemNode.ChildNodes.AddNode(), Aspose.Slides.SmartArt.SmartArtNode)

                    ' Adding text
                    newNode.TextFrame.Text = "New Node Added"

                End If
            Next shape

            ' Saving Presentation
            pres.Save(dataDir & "AddSmartArtNode_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx)
            ' ExEnd:AddNodes
        End Sub
    End Class
End Namespace