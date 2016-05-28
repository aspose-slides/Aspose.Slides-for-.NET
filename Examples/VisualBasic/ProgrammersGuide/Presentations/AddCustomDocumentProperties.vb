'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Imports System
Imports System.IO
Imports Aspose.Slides
Imports Aspose.Slides.Export

Namespace VisualBasic.Presentations
    Public Class AddCustomDocumentProperties
        Public Shared Sub Run()

            ' Loading a presentation
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            ' Instanciate the Presentation class
            Dim presentation As New Presentation()

            ' Getting Document Properties
            Dim documentProperties As IDocumentProperties = presentation.DocumentProperties

            ' Adding Custom properties
            documentProperties("New Custom") = 12
            documentProperties("My Name") = "Jawad"
            documentProperties("Custom") = 124

            ' Getting property name at particular index
            Dim getPropertyName As [String] = documentProperties.GetCustomPropertyName(2)

            ' Removing selected property
            documentProperties.RemoveCustomProperty(getPropertyName)

            ' Saving presentation
            presentation.Save(dataDir + "CustomDocumentProperties.pptx", SaveFormat.Pptx)

        End Sub
    End Class
End Namespace