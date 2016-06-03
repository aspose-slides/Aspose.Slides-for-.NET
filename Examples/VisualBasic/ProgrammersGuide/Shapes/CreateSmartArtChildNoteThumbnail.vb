'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Imports System
Imports System.Drawing
Imports System.Drawing.Imaging
Imports Aspose.Slides
Imports Aspose.Slides.SmartArt

Namespace VisualBasic.Shapes
    Public Class CreateSmartArtChildNoteThumbnail
        Public Shared Sub Run()

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            'Instantiate Presentation class that represents the PPTX file 
            Dim pres As New Presentation()

            'Add SmartArt 
            Dim smart As ISmartArt = pres.Slides(0).Shapes.AddSmartArt(10, 10, 400, 300, SmartArtLayoutType.BasicCycle)

            'Obtain the reference of a node by using its Index  
            Dim node As ISmartArtNode = smart.Nodes(1)

            'Get thumbnail
            Dim bmp As Bitmap = node.Shapes(0).GetThumbnail()

            'Save thumbnail
            bmp.Save(dataDir + "SmartArt_ChildNote_Thumbnail.jpeg", System.Drawing.Imaging.ImageFormat.Jpeg)

        End Sub
    End Class
End Namespace