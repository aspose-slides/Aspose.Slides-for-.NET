'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Imports System.Drawing
Imports Aspose.Slides
Imports Aspose.Slides.Export

Namespace Aspose.Slides.Examples.VisualBasic.Shapes
    Public Class ConnectShapeUsingConnectionSite
        Public Shared Sub Run()
			'ExStart:ConnectShapeUsingConnectionSite
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Instantiate Presentation class that represents the PPTX file
            Using presentation As New Presentation()
                ' Accessing shapes collection for selected slide
                Dim shapes As IShapeCollection = presentation.Slides(0).Shapes

                ' Adding connector shape to slide shape collection
                Dim connector As IConnector = shapes.AddConnector(ShapeType.BentConnector3, 0, 0, 10, 10)

                ' Add autoshape Ellipse
                Dim ellipse As IAutoShape = shapes.AddAutoShape(ShapeType.Ellipse, 0, 100, 100, 100)

                ' Add autoshape Rectangle
                Dim rectangle As IAutoShape = shapes.AddAutoShape(ShapeType.Rectangle, 100, 200, 100, 100)

                ' Joining Shapes to connectors
                connector.StartShapeConnectedTo = ellipse
                connector.EndShapeConnectedTo = rectangle

                ' Setting the desired connection site index of Ellipse shape for connector to get connected
                Dim wantedIndex As UInteger = 6

                ' Checking if desired index is less than maximum site index count
                If ellipse.ConnectionSiteCount > wantedIndex Then
                    ' Setting the desired connection site for connector on Ellipse
                    connector.StartShapeConnectionSiteIndex = wantedIndex
                End If

                ' Save presentation
                presentation.Save(dataDir + "ConnectShapeUsingConnectionSite_out.pptx", SaveFormat.Pptx)
            End Using
			'ExEnd:ConnectShapeUsingConnectionSite
        End Sub
    End Class
End Namespace










