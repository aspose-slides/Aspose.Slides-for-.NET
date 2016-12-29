Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports System.Drawing

Namespace Aspose.Slides.Examples.VisualBasic.Tables
    Public Class TableFromScratch
        Public Shared Sub Run()
            ' ExStart:TableFromScratch

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
                Dim sld As ISlide = pres.Slides(0)

                ' Define columns with widths and rows with heights
                Dim dblCols() As Double = {50, 50, 50}
                Dim dblRows() As Double = {50, 30, 30, 30, 30}

                ' Add table shape to slide
                Dim tbl As ITable = sld.Shapes.AddTable(100, 50, dblCols, dblRows)

                ' Set border format for each cell
                For Each row As IRow In tbl.Rows
                    For Each cell As ICell In row
                        cell.BorderTop.FillFormat.FillType = FillType.Solid
                        cell.BorderTop.FillFormat.SolidFillColor.Color = Color.Red
                        cell.BorderTop.Width = 5

                        cell.BorderBottom.FillFormat.FillType = FillType.Solid
                        cell.BorderBottom.FillFormat.SolidFillColor.Color = Color.Red
                        cell.BorderBottom.Width = 5

                        cell.BorderLeft.FillFormat.FillType = FillType.Solid
                        cell.BorderLeft.FillFormat.SolidFillColor.Color = Color.Red
                        cell.BorderLeft.Width = 5

                        cell.BorderRight.FillFormat.FillType = FillType.Solid
                        cell.BorderRight.FillFormat.SolidFillColor.Color = Color.Red
                        cell.BorderRight.Width = 5
                    Next cell
                Next row


                ' Merge cells 1 & 2 of row 1
                tbl.MergeCells(tbl(0, 0), tbl(1, 0), False)

                ' Add text to the merged cell
                tbl(0, 0).TextFrame.Text = "Merged Cells"

                'Write PPTX to Disk
                pres.Save(dataDir & "TableFromScratch_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx)
            End Using
            ' ExEnd:TableFromScratch
        End Sub
    End Class
End Namespace