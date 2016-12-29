Imports System
Imports System.Drawing
Imports System.Drawing.Imaging
Imports Aspose.Slides.Export
Imports Aspose.Slides.SmartArt
Imports Aspose.Slides

'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Namespace Aspose.Slides.Examples.VisualBasic.Shapes
    Public Class ConnectShapesUsingConnectors
        Public Shared Sub Run()
			'ExStart:ConnectShapesUsingConnectors
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Instantiate Presentation class that represents the PPTX file
            Using input As New Presentation()

                ' Accessing shapes collection for selected slide
                Dim shapes As IShapeCollection = input.Slides(0).Shapes

                ' Add autoshape Ellipse
                Dim ellipse As IAutoShape = shapes.AddAutoShape(ShapeType.Ellipse, 0, 100, 100, 100)

                ' Add autoshape Rectangle
                Dim rectangle As IAutoShape = shapes.AddAutoShape(ShapeType.Rectangle, 100, 300, 100, 100)

                ' Adding connector shape to slide shape collection
                Dim connector As IConnector = shapes.AddConnector(ShapeType.BentConnector2, 0, 0, 10, 10)

                ' Joining Shapes to connectors
                connector.StartShapeConnectedTo = ellipse
                connector.EndShapeConnectedTo = rectangle

                ' Call reroute to set the automatic shortest path between shapes
                connector.Reroute()

                ' Saving presenation
                input.Save(dataDir + "Connecting shapes using connectors_out.pptx", SaveFormat.Pptx)
            End Using
			'ExEnd:ConnectShapesUsingConnectors
        End Sub
    End Class
End Namespace


