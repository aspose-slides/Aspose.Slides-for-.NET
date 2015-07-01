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

Namespace VisualBasic.Shapes
    Public Class AccessOLEObjectFrame
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            'Load the PPTX to Presentation object
            Dim pres As New Presentation(dataDir & "AccessOLEObjectFrame.pptx")

            'Access the first slide
            Dim sld As ISlide = pres.Slides(0)

            'Cast the shape to OleObjectFrame
            Dim oof As OleObjectFrame = CType(sld.Shapes(0), OleObjectFrame)

            'Read the OLE Object and write it to disk
            If oof IsNot Nothing Then
                Dim fstr As New FileStream(dataDir & "excelFromOLE.xlsx", FileMode.Create, FileAccess.Write)
                Dim buf() As Byte = oof.ObjectData
                fstr.Write(buf, 0, buf.Length)
                fstr.Flush()
                fstr.Close()
            End If


        End Sub
    End Class
End Namespace