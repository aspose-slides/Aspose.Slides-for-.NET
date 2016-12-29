Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides

Namespace Aspose.Slides.Examples.VisualBasic.Tables
    Public Class TableWithCellBorders
        Public Shared Sub Run()
            ' ExStart:TableWithCellBorders

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Tables()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not IsExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If

            ' Instantiate Presentation class that represents PPTX file
            Using pres As New Presentation()

                ' Access first slide
                Dim sld As Slide = CType(pres.Slides(0), Slide)

                ' Define columns with widths and rows with heights
                Dim dblCols() As Double = {50, 50, 50, 50}
                Dim dblRows() As Double = {50, 30, 30, 30, 30}

                ' Add table shape to slide

                ' Add table shape to slide
                Dim tbl As ITable = sld.Shapes.AddTable(100, 50, dblCols, dblRows)

                ' Set border format for each cell
                For Each row As IRow In tbl.Rows
                    For Each cell As ICell In row
                        cell.BorderTop.FillFormat.FillType = FillType.NoFill
                        cell.BorderBottom.FillFormat.FillType = FillType.NoFill
                        cell.BorderLeft.FillFormat.FillType = FillType.NoFill
                        cell.BorderRight.FillFormat.FillType = FillType.NoFill
                    Next cell
                Next row

                'Write PPTX to Disk
                pres.Save(dataDir & "table1_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx)
            End Using
            ' ExEnd:TableWithCellBorders
        End Sub
    End Class
End Namespace