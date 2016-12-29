Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides

Namespace Aspose.Slides.Examples.VisualBasic.Tables
    Public Class RemovingRowColumn
        Public Shared Sub Run()
            ' ExStart:RemovingRowColumn

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Tables()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not IsExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If

            Dim pres As New Presentation()

            Dim slide As ISlide = pres.Slides(0)
            Dim colWidth() As Double = {100, 50, 30}
            Dim rowHeight() As Double = {30, 50, 30}

            Dim table As ITable = slide.Shapes.AddTable(100, 100, colWidth, rowHeight)

            table.Rows.RemoveAt(1, False)
            table.Columns.RemoveAt(1, False)
            pres.Save(dataDir & "TestTable_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx)
            ' ExEnd:RemovingRowColumn

        End Sub
    End Class
End Namespace