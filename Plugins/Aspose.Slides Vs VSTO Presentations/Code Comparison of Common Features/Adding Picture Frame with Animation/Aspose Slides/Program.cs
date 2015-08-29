using Aspose.Slides;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose_Slides
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creating empty presentation
            Presentation pres = new Presentation();

            //Accessing the First slide
            Slide slide = pres.GetSlideByPosition(1);

            //Adding the picture object to pictures collection of the presentation
            Picture pic = new Picture(pres, "pic.jpeg");

            //After the picture object is added, the picture is given a uniqe picture Id
            int picId = pres.Pictures.Add(pic);

            //Adding Picture Frame
            Shape PicFrame = slide.Shapes.AddPictureFrame(picId, 1450, 1100, 2500, 2200);

            //Applying animation on picture frame
            PicFrame.AnimationSettings.EntryEffect = ShapeEntryEffect.BoxIn;

            //Saving Presentation
            pres.Write("AsposeAnim.ppt");
        }
    }
}
