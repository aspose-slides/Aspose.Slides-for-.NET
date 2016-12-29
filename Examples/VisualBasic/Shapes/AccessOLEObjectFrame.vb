Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports System

Namespace Aspose.Slides.Examples.VisualBasic.Shapes
    Public Class AccessOLEObjectFrame
        Public Shared Sub Run()
			'ExStart:AccessOLEObjectFrame
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Load the PPTX to Presentation object
            Dim pres As New Presentation(dataDir & "AccessingOLEObjectFrame.pptx")

            ' Access the first slide
            Dim sld As ISlide = pres.Slides(0)

            ' Cast the shape to OleObjectFrame
            Dim oof As OleObjectFrame = CType(sld.Shapes(0), OleObjectFrame)

            ' Read the OLE Object and write it to disk
            If oof IsNot Nothing Then
                Dim fstr As New FileStream(dataDir & "excelFromOLE_out.xlsx", FileMode.Create, FileAccess.Write)
                Dim buf() As Byte = oof.ObjectData
                fstr.Write(buf, 0, buf.Length)
                fstr.Flush()
                fstr.Close()
            End If
			'ExEnd:AccessOLEObjectFrame
        End Sub
    End Class
End Namespace