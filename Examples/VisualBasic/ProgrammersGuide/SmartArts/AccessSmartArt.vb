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
Imports System

Namespace VisualBasic.SmartArts
    Public Class AccessSmartArt
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_SmartArts()

            'Load the desired the presentation
            'Load the desired the presentation
            Dim pres As New Presentation(dataDir & "AccessSmartArt.pptx")

            'Traverse through every shape inside first slide
            For Each shape As IShape In pres.Slides(0).Shapes

                'Check if shape is of SmartArt type
                If TypeOf shape Is SmartArt Then

                    'Typecast shape to SmartArt
                    Dim smart As SmartArt = CType(shape, SmartArt)

                    'Traverse through all nodes inside SmartArt
                    For i As Integer = 0 To smart.AllNodes.Count - 1
                        'Accessing SmartArt node at index i
                        Dim node As SmartArtNode = CType(smart.AllNodes(i), SmartArtNode)

                        'Printing the SmartArt node parameters
                        Dim outString As String = String.Format("i = {0}, Text = {1},  Level = {2}, Position = {3}", i, node.TextFrame.Text, node.Level, node.Position)
                        Console.WriteLine(outString)
                    Next i
                End If
            Next shape


        End Sub
    End Class
End Namespace