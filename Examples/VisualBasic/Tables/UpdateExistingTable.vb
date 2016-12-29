Imports Microsoft.VisualBasic
Imports System.IO
Imports Aspose.Slides

Namespace Aspose.Slides.Examples.VisualBasic.Tables
    Public Class UpdateExistingTable
        Public Shared Sub Run()
            ' ExStart:UpdateExistingTable
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Tables()

            ' Instantiate Presentation class that represents PPTX//Instantiate Presentation class that represents PPTX
            Using pres As New Presentation(dataDir & "UpdateExistingTable.pptx")

                ' Access the first slide
                Dim sld As ISlide = pres.Slides(0)

                ' Initialize null TableEx
                Dim tbl As ITable = Nothing

                ' Iterate through the shapes and set a reference to the table found
                For Each shp As IShape In sld.Shapes
                    If TypeOf shp Is ITable Then
                        tbl = CType(shp, ITable)
                    End If
                Next shp

                ' Set the text of the first column of second row
                tbl(0, 1).TextFrame.Text = "New"

                ' Write the PPTX to Disk
                pres.Save(dataDir & "table1_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx)
            End Using
            ' ExEnd:UpdateExistingTable
        End Sub
    End Class
End Namespace