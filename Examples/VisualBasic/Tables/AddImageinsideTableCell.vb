Imports System.Drawing
Imports Aspose.Slides.Export
Imports Aspose.Slides

'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Namespace Aspose.Slides.Examples.VisualBasic.Tables
    Public Class AddImageinsideTableCell
        Public Shared Sub Run()
            ' ExStart:AddImageinsideTableCell
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Tables()

            ' Instantiate Presentation class object
            Dim presentation As New Presentation()

            ' Access first slide
            Dim sld As ISlide = presentation.Slides(0)

            ' Define columns with widths and rows with heights
            Dim dblCols() As Double = {150, 150, 150, 150}
            Dim dblRows() As Double = {100, 100, 100, 100, 90}

            ' Add table shape to slide
            Dim tbl As ITable = sld.Shapes.AddTable(50, 50, dblCols, dblRows)

            ' Creating a Bitmap Image object to hold the image file
            Dim image As Bitmap = New Bitmap(dataDir + "aspose-logo.jpg")

            ' Create an IPPImage object using the bitmap object
            Dim imgx1 As IPPImage = presentation.Images.AddImage(image)

            ' Add image to first table cell
            tbl(0, 0).FillFormat.FillType = FillType.Picture
            tbl(0, 0).FillFormat.PictureFillFormat.PictureFillMode = PictureFillMode.Stretch
            tbl(0, 0).FillFormat.PictureFillFormat.Picture.Image = imgx1

            ' Save PPTX to Disk
            presentation.Save(dataDir + "Image_Inside_TableCell_out.pptx", SaveFormat.Pptx)
            ' ExStart:AddImageinsideTableCell
        End Sub
    End Class
End Namespace