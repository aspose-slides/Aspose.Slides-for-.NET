Imports System
Imports Aspose.Slides

'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx


Namespace Aspose.Slides.Examples.VisualBasic.Rendering.Printing
    Public Class SpecificPrinterPrinting
        Public Shared Sub Run()
			'ExStart:SpecificPrinterPrinting

            Try

                ' The path to the documents directory.
                Dim dataDir As String = RunExamples.GetDataDir_Rendering()

                ' Load the presentation
                Dim presentation As New Presentation(dataDir + "Print.ppt")

                ' Call the print method to print whole presentation to the desired printer
                presentation.Print("Please set your printer name here")
            Catch ex As Exception
                Console.WriteLine(ex.Message & "Please set printer name as string parameter to the Presentation Print method ")
            End Try
			'ExEnd:SpecificPrinterPrinting
        End Sub
    End Class
End Namespace