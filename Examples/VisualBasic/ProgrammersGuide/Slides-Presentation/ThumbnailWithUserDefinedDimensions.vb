'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Slides. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System.IO
Imports System.Drawing
Imports Aspose.Slides

'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Namespace VisualBasic.Slides
    Public Class ThumbnailWithUserDefinedDimensions
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations()

            'Instantiate a Presentation class that represents the presentation file
            Using pres As New Presentation(dataDir & "ThumbnailWithUserDefinedDimensions.pptx")

                'Access the first slide
                Dim sld As ISlide = pres.Slides(0)

                'User defined dimension
                Dim desiredX As Integer = 1200
                Dim desiredY As Integer = 800

                'Getting scaled value  of X and Y
                Dim ScaleX As Single = CSng(1.0 / pres.SlideSize.Size.Width) * desiredX
                Dim ScaleY As Single = CSng(1.0 / pres.SlideSize.Size.Height) * desiredY


                'Create a full scale image
                Dim bmp As Bitmap = sld.GetThumbnail(ScaleX, ScaleY)

                'Save the image to disk in JPEG format
                bmp.Save(dataDir & "Thumbnail2.jpg", System.Drawing.Imaging.ImageFormat.Jpeg)

            End Using
        End Sub


    End Class
End Namespace

