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
Imports Aspose.Slides.Export

Namespace VisualBasic.Shapes
    Public Class ChangeOLEObjectData
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            Dim pres As New Presentation(dataDir & "ChangeOLEObjectData.pptx")

            Dim slide As ISlide = pres.Slides(0)

            Dim ole As OleObjectFrame = Nothing

            'Traversing all shapes for Ole frame
            For Each shape As IShape In slide.Shapes
                If TypeOf shape Is OleObjectFrame Then
                    ole = CType(shape, OleObjectFrame)
                End If
            Next shape

            If ole IsNot Nothing Then
                'Reading object data in Workbook
                Dim Wb As Aspose.Cells.Workbook

                Using msln As New System.IO.MemoryStream(ole.ObjectData)
                    Wb = New Aspose.Cells.Workbook(msln)

                    Using msout As New System.IO.MemoryStream()
                        'Modifying the workbook data
                        Wb.Worksheets(0).Cells(0, 4).PutValue("E")
                        Wb.Worksheets(0).Cells(1, 4).PutValue(12)
                        Wb.Worksheets(0).Cells(2, 4).PutValue(14)
                        Wb.Worksheets(0).Cells(3, 4).PutValue(15)

                        Dim so1 As New Aspose.Cells.OoxmlSaveOptions(Aspose.Cells.SaveFormat.Xlsx)

                        Wb.Save(msout, so1)

                        'Changing Ole frame object data
                        msout.Position = 0
                        ole.ObjectData = msout.ToArray()
                    End Using
                End Using
            End If

            pres.Save(dataDir & "OleEdit.pptx", SaveFormat.Pptx)



        End Sub
    End Class
End Namespace