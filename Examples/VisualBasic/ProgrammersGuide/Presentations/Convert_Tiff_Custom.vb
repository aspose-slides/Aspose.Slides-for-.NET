'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Slides. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides.Export
Imports System.Drawing
Imports Aspose.Slides

'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Namespace VisualBasic.Presentations
    Public Class Convert_Tiff_Custom
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            'Instantiate a Presentation object that represents a Presentation file
            Using pres As New Presentation(dataDir & "Convert_Tiff_Custom.pptx")

                'Instantiate the TiffOptions class
                Dim opts As New Aspose.Slides.Export.TiffOptions()

                'Setting compression type
                opts.CompressionType = TiffCompressionTypes.Default

                'Compression Types

                'Default - Specifies the default compression scheme (LZW).
                'None - Specifies no compression.
                'CCITT3
                'CCITT4
                'LZW
                'RLE

                'Depth � depends on the compression type and cannot be set manually.
                'Resolution unit � is always equal to �2� (dots per inch)

                'Setting image DPI
                opts.DpiX = 200
                opts.DpiY = 100

                'Set Image Size
                opts.ImageSize = New Size(1728, 1078)

                'Save the presentation to TIFF with specified image size
                pres.Save(dataDir & "demo.tiff", Aspose.Slides.Export.SaveFormat.Tiff, opts)

            End Using
        End Sub
    End Class
End Namespace