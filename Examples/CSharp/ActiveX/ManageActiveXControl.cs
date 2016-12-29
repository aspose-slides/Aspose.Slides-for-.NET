using System;
using System.Drawing;
using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.ActiveX
{
    public class ManageActiveXControl
    {
        public static void Run()
        {
            //ExStart:ManageActiveXControl
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_ActiveX();

            // Accessing the presentation with  ActiveX controls
            Presentation presentation = new Presentation(dataDir + "ActiveX.pptm");

            // Accessing the first slide in presentation
            ISlide slide = presentation.Slides[0];

            // changing TextBox text
            IControl control = slide.Controls[0];

            if (control.Name == "TextBox1" && control.Properties != null)
            {
                string newText = "Changed text";
                control.Properties["Value"] = newText;

                // changing substitute image. Powerpoint will replace this image during activeX activation, so sometime it's OK to leave image unchanged.

                Bitmap image = new Bitmap((int)control.Frame.Width, (int)control.Frame.Height);
                Graphics graphics = Graphics.FromImage(image);
                Brush brush = new SolidBrush(Color.FromKnownColor(KnownColor.Window));
                graphics.FillRectangle(brush, 0, 0, image.Width, image.Height);
                brush.Dispose();
                System.Drawing.Font font = new System.Drawing.Font(control.Properties["FontName"], 14);
                brush = new SolidBrush(Color.FromKnownColor(KnownColor.WindowText));
                graphics.DrawString(newText, font, brush, 10, 4);
                brush.Dispose();
                Pen pen = new Pen(Color.FromKnownColor(KnownColor.ControlDark), 1);
                graphics.DrawLines(
                    pen, new System.Drawing.Point[] { new System.Drawing.Point(0, image.Height - 1), new System.Drawing.Point(0, 0), new System.Drawing.Point(image.Width - 1, 0) });
                pen.Dispose();
                pen = new Pen(Color.FromKnownColor(KnownColor.ControlDarkDark), 1);

                graphics.DrawLines(pen, new System.Drawing.Point[] { new System.Drawing.Point(1, image.Height - 2), new System.Drawing.Point(1, 1), new System.Drawing.Point(image.Width - 2, 1) });
                pen.Dispose();
                pen = new Pen(Color.FromKnownColor(KnownColor.ControlLight), 1);
                graphics.DrawLines(pen, new System.Drawing.Point[]
                {
                        new System.Drawing.Point(1, image.Height - 1), new System.Drawing.Point(image.Width - 1, image.Height - 1),
                        new System.Drawing.Point(image.Width - 1, 1)
                });
                pen.Dispose();
                pen = new Pen(Color.FromKnownColor(KnownColor.ControlLightLight), 1);
                graphics.DrawLines(pen,new System.Drawing.Point[] { new System.Drawing.Point(0, image.Height), new System.Drawing.Point(image.Width, image.Height), new System.Drawing.Point(image.Width, 0) });
                pen.Dispose();
                graphics.Dispose();
                control.SubstitutePictureFormat.Picture.Image = presentation.Images.AddImage(image);
            }

            // changing Button caption
            control = slide.Controls[1];

            if (control.Name == "CommandButton1" && control.Properties != null)
            {
                String newCaption = "MessageBox";
                control.Properties["Caption"] = newCaption;

                // changing substitute
                Bitmap image = new Bitmap((int)control.Frame.Width, (int)control.Frame.Height);
                Graphics graphics = Graphics.FromImage(image);
                Brush brush = new SolidBrush(Color.FromKnownColor(KnownColor.Control));
                graphics.FillRectangle(brush, 0, 0, image.Width, image.Height);
                brush.Dispose();
                System.Drawing.Font font = new System.Drawing.Font(control.Properties["FontName"], 14);
                brush = new SolidBrush(Color.FromKnownColor(KnownColor.WindowText));
                SizeF textSize = graphics.MeasureString(newCaption, font, int.MaxValue);
                graphics.DrawString(newCaption, font, brush, (image.Width - textSize.Width) / 2, (image.Height - textSize.Height) / 2);
                brush.Dispose();
                Pen pen = new Pen(Color.FromKnownColor(KnownColor.ControlLightLight), 1);
                graphics.DrawLines(pen, new System.Drawing.Point[] { new System.Drawing.Point(0, image.Height - 1), new System.Drawing.Point(0, 0), new System.Drawing.Point(image.Width - 1, 0) });
                pen.Dispose();
                pen = new Pen(Color.FromKnownColor(KnownColor.ControlLight), 1);
                graphics.DrawLines(pen, new System.Drawing.Point[] { new System.Drawing.Point(1, image.Height - 2), new System.Drawing.Point(1, 1), new System.Drawing.Point(image.Width - 2, 1) });
                pen.Dispose();
                pen = new Pen(Color.FromKnownColor(KnownColor.ControlDark), 1);
                graphics.DrawLines(pen,new System.Drawing.Point[]
                {
                    new System.Drawing.Point(1, image.Height - 1),
                    new System.Drawing.Point(image.Width - 1, image.Height - 1),
                    new System.Drawing.Point(image.Width - 1, 1)
                });
                pen.Dispose();
                pen = new Pen(Color.FromKnownColor(KnownColor.ControlDarkDark), 1);
                graphics.DrawLines(pen,new System.Drawing.Point[] { new System.Drawing.Point(0, image.Height), new System.Drawing.Point(image.Width, image.Height), new System.Drawing.Point(image.Width, 0) });
                pen.Dispose();
                graphics.Dispose();
                control.SubstitutePictureFormat.Picture.Image = presentation.Images.AddImage(image);
            }

            // Moving ActiveX frames 100 points down
            foreach (Control ctl in slide.Controls)
            {
                IShapeFrame frame = control.Frame;
                control.Frame = new ShapeFrame(
                    frame.X, frame.Y + 100, frame.Width, frame.Height, frame.FlipH, frame.FlipV, frame.Rotation);
            }

            // Save the presentation with Edited ActiveX Controls
            presentation.Save(dataDir + "withActiveX-edited_out.pptm", Aspose.Slides.Export.SaveFormat.Pptm);


            // Now removing controls
            slide.Controls.Clear();

            // Saving the presentation with cleared ActiveX controls
            presentation.Save(dataDir + "withActiveX.cleared_out.pptm", Aspose.Slides.Export.SaveFormat.Pptm);
            //ExEnd:ManageActiveXControl
        }
    }
}