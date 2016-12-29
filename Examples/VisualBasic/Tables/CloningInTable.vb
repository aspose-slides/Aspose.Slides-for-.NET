Imports System.Drawing
Imports Aspose.Slides.Export
Imports Aspose.Slides

'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Namespace Aspose.Slides.Examples.VisualBasic.Tables
    Public Class CloningInTable
        Public Shared Sub Run()

            ' ExStart:CloningInTable
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Tables()

            ' Creating empty presentation
            Dim presentation As New Presentation()

            ' Access first slide
            Dim sld As ISlide = presentation.Slides(0)

            ' Define columns with widths and rows with heights
            Dim dblCols As Double() = {50, 50, 50}
            Dim dblRows As Double() = {50, 30, 30, 30}

            ' Add table shape to slide
            Dim table As ITable = sld.Shapes.AddTable(100, 50, dblCols, dblRows)

            ' Set border format for each cell
            For Each row As IRow In table.Rows
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

            table(0, 0).TextFrame.Text = "00"
            table(0, 1).TextFrame.Text = "01"
            table(0, 2).TextFrame.Text = "02"
            table(0, 3).TextFrame.Text = "03"
            table(1, 0).TextFrame.Text = "10"
            table(2, 0).TextFrame.Text = "20"
            table(1, 1).TextFrame.Text = "11"
            table(2, 1).TextFrame.Text = "21"

            ' AddClone adds a row in the end of the table
            table.Rows.AddClone(table.Rows(0), False)

            ' InsertClone adds a row at specific position in a table
            table.Rows.InsertClone(2, table.Rows(0), False)

            ' AddClone adds a column in the end of the table
            table.Columns.AddClone(table.Columns(0), False)

            ' InsertClone adds a column at specific position in a table
            table.Columns.InsertClone(2, table.Columns(0), False)

            presentation.Save(dataDir + "CloneRow_out.pptx", SaveFormat.Pptx)
            ' ExEnd:CloningInTable
        End Sub
    End Class
End Namespace