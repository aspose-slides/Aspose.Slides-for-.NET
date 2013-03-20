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

Namespace ConvertPPTXToXPS
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate a PresentationEx object that represents a PPTX file
			Dim pres As New PresentationEx(dataDir & "demo.pptx")


			' 1.
			'Saving the presentation to TIFF document
			pres.Save(dataDir & "output.xps", Aspose.Slides.Export.SaveFormat.Xps)


			' 2.
			'Instantiate the TiffOptions class
			Dim opts As New Aspose.Slides.Export.XpsOptions()

			'Save MetaFiles as PNG
			opts.SaveMetafilesAsPng = True

			'Save the presentation to XPS document
			pres.Save(dataDir & "outputWithXPSOptions.xps", Aspose.Slides.Export.SaveFormat.Xps, opts)
		End Sub
	End Class
End Namespace