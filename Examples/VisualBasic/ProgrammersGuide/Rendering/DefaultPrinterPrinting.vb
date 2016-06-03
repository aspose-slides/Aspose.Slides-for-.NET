'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Imports System
Imports System.IO
Imports Aspose.Slides
Imports Aspose.Slides.Export

Namespace VisualBasic.Rendering
    Public Class DefaultPrinterPrinting
        Public Shared Sub Run()

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Rendering()

            ' Load the presentation
            Dim presentation As New Presentation(dataDir + "Print.ppt")

            ' Call the print method to print whole presentation to the default printer
            presentation.Print()

        End Sub
    End Class
End Namespace