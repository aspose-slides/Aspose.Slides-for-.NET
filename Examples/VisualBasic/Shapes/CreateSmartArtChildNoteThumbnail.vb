Imports System.Drawing
Imports Aspose.Slides.SmartArt
Imports Aspose.Slides

'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Namespace Aspose.Slides.Examples.VisualBasic.Shapes
    Public Class CreateSmartArtChildNoteThumbnail
        Public Shared Sub Run()
			'ExStart:CreateSmartArtChildNoteThumbnail
            ' For complete examples and data files, please go to https://github.com/aspose-slides/Aspose.Slides-for-.NET

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Instantiate Presentation class that represents the PPTX file 
            Dim presentation As New Presentation()

            ' Add SmartArt 
            Dim smart As ISmartArt = presentation.Slides(0).Shapes.AddSmartArt(10, 10, 400, 300, SmartArtLayoutType.BasicCycle)

            ' Obtain the reference of a node by using its Index  
            Dim node As ISmartArtNode = smart.Nodes(1)

            ' Get thumbnail
            Dim bmp As Bitmap = node.Shapes(0).GetThumbnail()

            ' Save thumbnail
            bmp.Save(dataDir + "SmartArt_ChildNote_Thumbnail_out.jpeg", System.Drawing.Imaging.ImageFormat.Jpeg)
			'ExEnd:CreateSmartArtChildNoteThumbnail
        End Sub
    End Class
End Namespace