Imports Microsoft.VisualBasic
Imports System.IO
Imports Aspose.Slides
Imports AsposeSlides = Aspose.Slides.SmartArt

Namespace Aspose.Slides.Examples.VisualBasic.SmartArts
    Public Class AddNodes
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_SmartArts()

            'Load the desired the presentation//Load the desired the presentation
            Dim pres As New Presentation(dataDir & "AddNodes.pptx")

            'Traverse through every shape inside first slide
            For Each shape As IShape In pres.Slides(0).Shapes

                'Check if shape is of SmartArt type
                If TypeOf shape Is AsposeSlides.SmartArt Then

                    'Typecast shape to SmartArt
                    Dim smart As AsposeSlides.SmartArt = CType(shape, AsposeSlides.SmartArt)

                    'Adding a new SmartArt Node
                    Dim TemNode As AsposeSlides.SmartArtNode = CType(smart.AllNodes.AddNode(), AsposeSlides.SmartArtNode)

                    'Adding text
                    TemNode.TextFrame.Text = "Test"

                    'Adding new child node in parent node. It  will be added in the end of collection
                    Dim newNode As AsposeSlides.SmartArtNode = CType(TemNode.ChildNodes.AddNode(), AsposeSlides.SmartArtNode)

                    'Adding text
                    newNode.TextFrame.Text = "New Node Added"

                End If
            Next shape

            'Saving Presentation
            pres.Save(dataDir & "AddSmartArtNode.pptx", Aspose.Slides.Export.SaveFormat.Pptx)


        End Sub
    End Class
End Namespace