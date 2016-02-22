using Aspose.Slides;
using Aspose.Slides.Pptx;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rendering_Individual_Shapes_as_Images
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = @"Files\";
            //Instantiate a Presentation object that represents a PPT file
            Presentation pres = new Presentation(path + "RenderShapeAsImage.ppt");

            //Accessing a slide using its slide position
            Slide slide = pres.GetSlideByPosition(2);


            //Iterate all shapes on a slide and create thumbnails
            ShapeCollection shapes = slide.Shapes;
            for (int i = 0; i < shapes.Count; i++)
            {
                Shape shape = shapes[i];
                //Getting the thumbnail image of the shape
                Image img = slide.GetThumbnail(new object[] { shape }, 1.0, 1.0,shape.ShapeRectangle);
                //Saving the thumbnail image in gif format
                img.Save(path + i + ".gif", ImageFormat.Gif);
            }
        }
    }           
}
