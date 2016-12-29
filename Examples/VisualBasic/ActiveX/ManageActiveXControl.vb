Imports System
Imports System.Drawing
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.ActiveX
    Public Class ManageActiveXControl
        Public Shared Sub Run()
        	'ExStart:ManageActiveXControl
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_ActiveX()

            ' Accessing the presentation with  ActiveX controls
            Dim presentation As New Presentation(dataDir & Convert.ToString("ActiveX.pptm"))

            ' Accessing the first slide in presentation
            Dim slide As ISlide = presentation.Slides(0)

            ' changing TextBox text
            Dim control As IControl = slide.Controls(0)


            If control.Name = "TextBox1" AndAlso control.Properties IsNot Nothing Then
                Dim newText As String = "Changed text"
                control.Properties("Value") = newText

                ' changing substitute image. Powerpoint will replace this image during activeX activation, so sometime it' S OK to leave image unchanged.

                Dim image As New Bitmap(CInt(control.Frame.Width), CInt(control.Frame.Height))
                Dim graphics__1 As Graphics = Graphics.FromImage(image)
                Dim brush As Brush = New SolidBrush(Color.FromKnownColor(KnownColor.Window))
                graphics__1.FillRectangle(brush, 0, 0, image.Width, image.Height)
                brush.Dispose()
                Dim font As New System.Drawing.Font(control.Properties("FontName"), 14)
                brush = New SolidBrush(Color.FromKnownColor(KnownColor.WindowText))
                graphics__1.DrawString(newText, font, brush, 10, 4)
                brush.Dispose()
                Dim pen As New Pen(Color.FromKnownColor(KnownColor.ControlDark), 1)
                graphics__1.DrawLines(pen, New System.Drawing.Point() {New System.Drawing.Point(0, image.Height - 1), New System.Drawing.Point(0, 0), New System.Drawing.Point(image.Width - 1, 0)})
                pen.Dispose()
                pen = New Pen(Color.FromKnownColor(KnownColor.ControlDarkDark), 1)

                graphics__1.DrawLines(pen, New System.Drawing.Point() {New System.Drawing.Point(1, image.Height - 2), New System.Drawing.Point(1, 1), New System.Drawing.Point(image.Width - 2, 1)})
                pen.Dispose()
                pen = New Pen(Color.FromKnownColor(KnownColor.ControlLight), 1)
                graphics__1.DrawLines(pen, New System.Drawing.Point() {New System.Drawing.Point(1, image.Height - 1), New System.Drawing.Point(image.Width - 1, image.Height - 1), New System.Drawing.Point(image.Width - 1, 1)})
                pen.Dispose()
                pen = New Pen(Color.FromKnownColor(KnownColor.ControlLightLight), 1)
                graphics__1.DrawLines(pen, New System.Drawing.Point() {New System.Drawing.Point(0, image.Height), New System.Drawing.Point(image.Width, image.Height), New System.Drawing.Point(image.Width, 0)})
                pen.Dispose()
                graphics__1.Dispose()
                control.SubstitutePictureFormat.Picture.Image = presentation.Images.AddImage(image)
            End If

            ' changing Button caption
            control = slide.Controls(1)

            If control.Name = "CommandButton1" AndAlso control.Properties IsNot Nothing Then
                Dim newCaption As [String] = "MessageBox"
                control.Properties("Caption") = newCaption

                ' changing substitute
                Dim image As New Bitmap(CInt(control.Frame.Width), CInt(control.Frame.Height))
                Dim graphics__1 As Graphics = Graphics.FromImage(image)
                Dim brush As Brush = New SolidBrush(Color.FromKnownColor(KnownColor.Control))
                graphics__1.FillRectangle(brush, 0, 0, image.Width, image.Height)
                brush.Dispose()
                Dim font As New System.Drawing.Font(control.Properties("FontName"), 14)
                brush = New SolidBrush(Color.FromKnownColor(KnownColor.WindowText))
                Dim textSize As SizeF = graphics__1.MeasureString(newCaption, font, Integer.MaxValue)
                graphics__1.DrawString(newCaption, font, brush, (image.Width - textSize.Width) / 2, (image.Height - textSize.Height) / 2)
                brush.Dispose()
                Dim pen As New Pen(Color.FromKnownColor(KnownColor.ControlLightLight), 1)
                graphics__1.DrawLines(pen, New System.Drawing.Point() {New System.Drawing.Point(0, image.Height - 1), New System.Drawing.Point(0, 0), New System.Drawing.Point(image.Width - 1, 0)})
                pen.Dispose()
                pen = New Pen(Color.FromKnownColor(KnownColor.ControlLight), 1)
                graphics__1.DrawLines(pen, New System.Drawing.Point() {New System.Drawing.Point(1, image.Height - 2), New System.Drawing.Point(1, 1), New System.Drawing.Point(image.Width - 2, 1)})
                pen.Dispose()
                pen = New Pen(Color.FromKnownColor(KnownColor.ControlDark), 1)
                graphics__1.DrawLines(pen, New System.Drawing.Point() {New System.Drawing.Point(1, image.Height - 1), New System.Drawing.Point(image.Width - 1, image.Height - 1), New System.Drawing.Point(image.Width - 1, 1)})
                pen.Dispose()
                pen = New Pen(Color.FromKnownColor(KnownColor.ControlDarkDark), 1)
                graphics__1.DrawLines(pen, New System.Drawing.Point() {New System.Drawing.Point(0, image.Height), New System.Drawing.Point(image.Width, image.Height), New System.Drawing.Point(image.Width, 0)})
                pen.Dispose()
                graphics__1.Dispose()
                control.SubstitutePictureFormat.Picture.Image = presentation.Images.AddImage(image)
            End If

            ' Moving ActiveX frames 100 points down
            For Each ctl As Control In slide.Controls
                Dim frame As IShapeFrame = control.Frame
                control.Frame = New ShapeFrame(frame.X, frame.Y + 100, frame.Width, frame.Height, frame.FlipH, frame.FlipV, _
                    frame.Rotation)
            Next

            ' Save the presentation with Edited ActiveX Controls
            presentation.Save(dataDir & Convert.ToString("withActiveX-edited_out.pptm"), Aspose.Slides.Export.SaveFormat.Pptm)

            ' Now removing controls
            slide.Controls.Clear()

            ' Saving the presentation with cleared ActiveX controls
            presentation.Save(dataDir & Convert.ToString("withActiveX.cleared_out.pptm"), Aspose.Slides.Export.SaveFormat.Pptm)
        	'ExEnd:ManageActiveXControl
        End Sub
    End Class
End Namespace