'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Slides. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////
Imports System.IO

Imports Aspose.Slides
Imports Aspose.Slides.Pptx

Namespace ConvertingPPTXtoPDF
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' 1.
			' Conversion of PDF using default options.

			'Instantiate a PresentationEx object that represents a PPTX file
			Dim pres As New PresentationEx(dataDir & "demo.pptx")

			'Saving the PPTX presentation to PDF document
			pres.Save(dataDir & "demo1.pdf", Aspose.Slides.Export.SaveFormat.Pdf)

			' Display result of conversion.
			System.Console.WriteLine("Conversion to PDF performed successfully with default options!")

			' 2.
			' Conversion to PDF using custom options.

			'Instantiate the PdfOptions class
			Dim opts As New Aspose.Slides.Export.PdfOptions()

			'Set Jpeg Quality
			opts.JpegQuality = 90

			'Define behavior for meta files
			opts.SaveMetafilesAsPng = True

			'Set Text Compression level
			opts.TextCompression = Aspose.Slides.Export.PdfTextCompression.Flate

			'Define the PDF standard
			opts.Compliance = Aspose.Slides.Export.PdfCompliance.Pdf15

			'Save the presentation to PDF with specified options
			pres.Save(dataDir & "demo2.pdf", Aspose.Slides.Export.SaveFormat.Pdf, opts)

			' Display result of conversion.
			System.Console.WriteLine("Conversion to PDF performed successfully with custom options!")
		End Sub
	End Class
End Namespace