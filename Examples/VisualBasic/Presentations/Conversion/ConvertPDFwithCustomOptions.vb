Imports Aspose.Slides.Export
Imports Aspose.Slides

'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Namespace Aspose.Slides.Examples.VisualBasic.Presentations
    Public Class ConvertPDFwithCustomOptions
        Public Shared Sub Run()
			'ExStart:ConvertPDFwithCustomOptions
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            ' Instantiate a Presentation object that represents a presentation file
            Dim presentation As New Presentation(dataDir & "ConvertToPDF.pptx")

            ' Instantiate the PdfOptions class
            Dim pdfOptions As New Export.PdfOptions()

            ' Set Jpeg Quality
            pdfOptions.JpegQuality = 90

            ' Define behavior for metafiles
            pdfOptions.SaveMetafilesAsPng = True

            ' Set Text Compression level
            pdfOptions.TextCompression = Export.PdfTextCompression.Flate

            ' Define the PDF standard
            pdfOptions.Compliance = Export.PdfCompliance.Pdf15

            ' Save the presentation to PDF with specified options
            presentation.Save(dataDir & "Custom_Option_Pdf_Conversion_out.pdf", SaveFormat.Pdf, pdfOptions)
			'ExEnd:ConvertPDFwithCustomOptions
        End Sub
    End Class
End Namespace