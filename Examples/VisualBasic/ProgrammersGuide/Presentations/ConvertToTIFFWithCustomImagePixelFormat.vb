'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
Imports Aspose.Slides
Imports Aspose.Slides.Export
Imports VisualBasic

Namespace ProgrammersGuide.Presentations
    Public Class ConvertToTIFFWithCustomImagePixelFormat
        Public Shared Sub Run()

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            'Instantiate a Presentation object that represents a presentation file
            Using presentation As Presentation = New Presentation(dataDir + "DemoFile.pptx")

                'Instantiate the TiffOptions class
                Dim options As TiffOptions = New TiffOptions()
                'Setting Pixel Format
                options.PixelFormat = ImagePixelFormat.Format8bppIndexed

                'ImagePixelFormat contains the following values (as could be seen from documentation):
                'Format1bppIndexed; // 1 bits per pixel, indexed.
                'Format4bppIndexed; // 4 bits per pixel, indexed.
                'Format8bppIndexed; // 8 bits per pixel, indexed.
                'Format24bppRgb; // 24 bits per pixel, RGB.
                'Format32bppArgb; // 32 bits per pixel, ARGB.

                'Save the presentation to TIFF with specified image size
                presentation.Save(dataDir & "Tiff_With_Custom_Image_Pixel_Format.tiff", SaveFormat.Tiff, options)
            End Using

        End Sub
    End Class
End Namespace