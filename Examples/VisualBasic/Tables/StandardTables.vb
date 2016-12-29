Imports System
Imports System.Drawing
Imports Aspose.Slides.Export
Imports Aspose.Slides

'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Namespace Aspose.Slides.Examples.VisualBasic.Tables
    Public Class StandardTables
        Public Shared Sub Run()
            ' ExStart:StandardTables
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Tables()

            ' Instantiate Presentation class that represents PPTX file
            Using pres As New Presentation()

                ' Access first slide
                Dim sld As ISlide = pres.Slides(0)

                ' Define columns with widths and rows with heights
                Dim dblCols As Double() = {70, 70, 70, 70}
                Dim dblRows As Double() = {70, 70, 70, 70}

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
                    Next
                Next
                ' Write PPTX to Disk
                pres.Save(dataDir & Convert.ToString("StandardTables_out.pptx"), SaveFormat.Pptx)
            End Using
            ' ExEnd:StandardTables
        End Sub
    End Class
End Namespace
