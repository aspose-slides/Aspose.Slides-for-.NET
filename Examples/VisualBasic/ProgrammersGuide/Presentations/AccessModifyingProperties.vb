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
Imports Aspose.Slides

'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx


Namespace VisualBasic.Presentations
    Public Class AccessModifyingProperties
        Public Shared Sub Run()

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            'Instanciate the Presentation class that represents the PPTX
            Dim pres As New Presentation(dataDir & "AccessModifyingProperties.pptx")

            'Create a reference to DocumentProperties object associated with Prsentation
            Dim dp As IDocumentProperties = pres.DocumentProperties


            'Access and modify custom properties
            For i As Integer = 0 To dp.CountOfCustomProperties - 1
                'Display names and values of custom properties
                System.Console.WriteLine("Custom Property Name : " & dp.GetCustomPropertyName(i))
                System.Console.WriteLine("Custom Property Value : " & dp.GetCustomPropertyName(i))

                'Modify values of custom properties
                dp(dp.GetCustomPropertyName(i)) = "New Value " & (i + 1)
            Next i

            'Save your presentation to a file
            pres.Save(dataDir & "CustomDemoModified.pptx", SaveFormat.Pptx)

        End Sub
    End Class
End Namespace