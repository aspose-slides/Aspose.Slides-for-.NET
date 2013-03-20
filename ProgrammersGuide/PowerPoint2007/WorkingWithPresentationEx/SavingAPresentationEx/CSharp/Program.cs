//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Pptx;
using System.Drawing;

namespace SavingAPresentationEx
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");



            // 1.
            // Save presentation to file.

            //Instantiate a Presentation object that represents a PPT file
            PresentationEx pres1 = new PresentationEx();

            //...do some work here...
            ProcessPresentation(pres1);
            
            //Save your presentation to a file
            pres1.Write(dataDir + "toFile.pptx");


            
            // 2. 
            // Save your presentation to a stream

            //Instantiate a Presentation object that represents a PPT file
            PresentationEx pres2 = new PresentationEx();

            //...do some work here...
            ProcessPresentation(pres2);

            //Accessing the output stream of Http Response
            System.IO.Stream st = new FileStream(dataDir + "toStream.pptx", FileMode.OpenOrCreate);

            //Saving the presentation to the output stream of Http Response
            pres2.Write(st);

            // Close the stream.
            st.Close();


            
            // 3.
            // Saving a presentation with password protection.

            //Instantiate a Presentation object that represents a PPT file
            PresentationEx pres3 = new PresentationEx();

            //...do some work here...
            ProcessPresentation(pres3);

            //Setting Password
            pres3.Encrypt("test");

            //Save your presentation to a file
            pres3.Write(dataDir + "passwordProtected.pptx");



            // 4.
            // Save password protected Presentation with Read Access to Document Properties

            //Instantiate a Presentation object that represents a PPT file
            PresentationEx pres4 = new PresentationEx();

            //...do some work here...
            ProcessPresentation(pres4);

            //Setting access to document properties in password protected mode
            pres4.EncryptDocumentProperties = false;

            //Setting Password
            pres4.Encrypt("test");

            //Save your presentation to a file
            pres4.Write(dataDir + "passwordProtectedReadOnlyProperties.pptx");


            
            // 5.
            // Save a read only presentation.

            //Instantiate a Presentation object that represents a PPT file
            PresentationEx pres5 = new PresentationEx();

            //...do some work here...
            ProcessPresentation(pres5);

            //Setting Write protection Password
            pres5.SetWriteProtection("test");

            //Save your presentation to a file
            pres5.Write(dataDir + "readOnlyPresentation.pptx");



            // 6.
            // Removing Write Protection from a Presentation

            //Opening the presentation file
            PresentationEx pres6 = new PresentationEx(dataDir + "readOnlyPresentation.pptx");

            //Checking if presentation is write protected
            if(pres6.IsWriteProtected)
	            //Removing Write protection	
	            pres6.RemoveWriteProtection();
	
            //Saving presentation
            pres6.Write(dataDir + "writeProtectionRemoved.pptx");
        }

        public static void ProcessPresentation (PresentationEx pres)
        {
            //Get the first slide
            SlideEx sld = pres.Slides[0];

            //Add an AutoShape of Rectangle type
            int idx = sld.Shapes.AddAutoShape(ShapeTypeEx.Rectangle, 150, 75, 150, 50);
            AutoShapeEx ashp = (AutoShapeEx)sld.Shapes[idx];

            //Add TextFrame to the Rectangle
            ashp.AddTextFrame("Aspose");

            //Change the text color to Black (which is White by default)
            ashp.TextFrame.Paragraphs[0].Portions[0].FillFormat.FillType = FillTypeEx.Solid;
            ashp.TextFrame.Paragraphs[0].Portions[0].FillFormat.SolidFillColor.Color = Color.Black;

            //Change the line color of the rectangle to White
            ashp.ShapeStyle.LineColor.Color = System.Drawing.Color.White;

            //Remove any fill formatting in the shape
            ashp.FillFormat.FillType = FillTypeEx.NoFill;
        }
    }
}