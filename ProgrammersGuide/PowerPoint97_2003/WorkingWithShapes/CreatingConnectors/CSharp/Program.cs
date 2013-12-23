//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;
using System.Drawing;

namespace CreatingConnectors
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            //Instantiate a Presentation object with new empty PPT file
            Presentation pres = new Presentation();


            //Accessing a slide using its slide position
            Slide slide = pres.GetSlideByPosition(1);


            //Creating 4 rectangles
            Aspose.Slides.Rectangle root = CreateRectangle(slide, 500, 500, 2760, 500, "Connectors");
            Aspose.Slides.Rectangle straight = CreateRectangle(slide, 200, 3500, 2000, 400, "Straight");
            Aspose.Slides.Rectangle elbow = CreateRectangle(slide, 3500, 1500, 2000, 400, "Elbow");
            Aspose.Slides.Rectangle curve = CreateRectangle(slide, 3000, 2500, 2000, 400, "Curve");

            //Create straight connector
            CreateConnector(slide, ConnectorType.Straight, root, 2, straight, 0);


            //Create elbow connector
            CreateConnector(slide, ConnectorType.Elbow, root, 3, elbow, 0);


            //Create curve connector
            CreateConnector(slide, ConnectorType.Curve, root, 2, curve, 1);

            pres.Write(dataDir + "output.ppt");

        }

        static Aspose.Slides.Rectangle CreateRectangle(Slide slide, int x, int y, int w, int h, string text)
        {
            // Create new Rectangle shape on a slide
            Aspose.Slides.Rectangle rectangle = slide.Shapes.AddRectangle(x, y, w, h);


            // Set format of lines for the rectangle
            rectangle.LineFormat.Width = 5;
            rectangle.LineFormat.ForeColor = Color.Red;


            // Add centered text
            rectangle.AddTextFrame(text);
            TextFrame tf = rectangle.TextFrame;
            tf.Paragraphs[0].Alignment = TextAlignment.Center;
            tf.Paragraphs[0].Portions[0].FontBold = true;
            tf.Paragraphs[0].Portions[0].FontHeight = 36;


            // Return created shape
            return rectangle;
        }


        static Connector CreateConnector(Slide slide, ConnectorType type,
                                         Shape shape1, int connPoint1,
                                         Shape shape2, int connPoint2)
        {
            // Add new connector with some random default coordinates
            Connector connector = slide.Shapes.AddConnector(
                type, new Point(500, 500), new Point(1000, 1000));


            // Connect connector with 2 shapes
            connector.ConnectBegin(shape1, connPoint1);
            connector.ConnectEnd(shape2, connPoint2);


            // Set connector style
            connector.LineFormat.ForeColor = Color.Blue;
            connector.LineFormat.Width = 5;
            connector.LineFormat.EndArrowheadStyle = LineArrowheadStyle.Open;


            // Return created connector
            return connector;
        }


    }
}