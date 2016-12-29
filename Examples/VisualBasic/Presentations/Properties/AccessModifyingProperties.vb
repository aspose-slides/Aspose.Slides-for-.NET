Imports System
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Presentations.Properties
    Public Class AccessModifyingProperties
        Public Shared Sub Run()
			'ExStart:AccessModifyingProperties
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_PresentationProperties()

            ' Instanciate the Presentation class that represents the PPTX
            Dim presentation As New Presentation(dataDir & Convert.ToString("AccessModifyingProperties.pptx"))

            ' Create a reference to DocumentProperties object associated with Prsentation
            Dim documentProperties As IDocumentProperties = presentation.DocumentProperties

            ' Access and modify custom properties
            For i As Integer = 0 To documentProperties.CountOfCustomProperties - 1
                ' Display names and values of custom properties
                System.Console.WriteLine("Custom Property Name : " + documentProperties.GetCustomPropertyName(i).ToString())
                System.Console.WriteLine("Custom Property Value : " + documentProperties(documentProperties.GetCustomPropertyName(i)).ToString())

                ' Modify values of custom properties
                documentProperties(documentProperties.GetCustomPropertyName(i)) = "New Value" & (i + 1)
            Next

            ' Save your presentation to a file
            presentation.Save(dataDir & Convert.ToString("CustomDemoModified_out.pptx"), Aspose.Slides.Export.SaveFormat.Pptx)
			'ExEnd:AccessModifyingProperties
        End Sub
    End Class
End Namespace
