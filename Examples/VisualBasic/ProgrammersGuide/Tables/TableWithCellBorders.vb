'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Slides. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System.IO
Imports Aspose.Slides

'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Namespace VisualBasic.Tables
    Public Class TableWithCellBorders
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Tables()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not IsExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If

            'Instantiate Presentation class that represents PPTX file
            Using pres As New Presentation()

                'Access first slide
                Dim sld As Slide = CType(pres.Slides(0), Slide)

                'Define columns with widths and rows with heights
                Dim dblCols() As Double = {50, 50, 50, 50}
                Dim dblRows() As Double = {50, 30, 30, 30, 30}

                'Add table shape to slide

                'Add table shape to slide
                Dim tbl As ITable = sld.Shapes.AddTable(100, 50, dblCols, dblRows)

                'Set border format for each cell
                For Each row As IRow In tbl.Rows
                    For Each cell As ICell In row
                        cell.BorderTop.FillFormat.FillType = FillType.NoFill
                        cell.BorderBottom.FillFormat.FillType = FillType.NoFill
                        cell.BorderLeft.FillFormat.FillType = FillType.NoFill
                        cell.BorderRight.FillFormat.FillType = FillType.NoFill
                    Next cell
                Next row



                'Write PPTX to Disk
                pres.Save(dataDir & "table.pptx", Aspose.Slides.Export.SaveFormat.Pptx)
            End Using

        End Sub
    End Class
End Namespace