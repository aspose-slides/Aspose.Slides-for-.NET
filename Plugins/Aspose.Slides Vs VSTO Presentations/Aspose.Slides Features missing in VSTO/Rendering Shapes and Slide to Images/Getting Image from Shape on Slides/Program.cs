using Aspose.Slides.Pptx;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Getting_Image_from_Shape_on_Slides
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = @"Files\";
            //Accessing the presentation
            PresentationEx pres = new PresentationEx(path + "RenderImageFromShape.pptx");
            ImageEx img = null;
            int slideIndex = 0;
            String ImageType = "";
            bool ifImageFound = false;
            for (int i = 0; i < pres.Slides.Count; i++)
            {
                slideIndex++;
                //Accessing the first slide
                SlideEx sl = pres.Slides[i];
                System.Drawing.Imaging.ImageFormat Format = System.Drawing.Imaging.ImageFormat.Jpeg;
                for (int j = 0; j < sl.Shapes.Count; j++)
                {
                    // Accessing the shape with picture
                    ShapeEx sh = sl.Shapes[j];

                    if (sh is AutoShapeEx)
                    {
                        AutoShapeEx ashp = (AutoShapeEx)sh;
                        if (ashp.FillFormat.FillType == FillTypeEx.Picture)
                        {
                            img = ashp.FillFormat.PictureFillFormat.Picture.Image;
                            ImageType = img.ContentType;
                            ImageType = ImageType.Remove(0, ImageType.IndexOf("/") + 1);
                            ifImageFound = true;

                        }
                    }

                    else if (sh is PictureFrameEx)
                    {
                        PictureFrameEx pf = (PictureFrameEx)sh;
                        if (pf.FillFormat.FillType == FillTypeEx.Picture)
                        {
                            img = pf.PictureFormat.Picture.Image;
                            ImageType = img.ContentType;
                            ImageType = ImageType.Remove(0, ImageType.IndexOf("/") + 1);
                            ifImageFound = true;
                        }
                    }


                    //
                    //Setting the desired picture format
                    if (ifImageFound)
                    {
                        switch (ImageType)
                        {
                            case "jpeg":
                                Format = System.Drawing.Imaging.ImageFormat.Jpeg;
                                break;

                            case "emf":
                                Format = System.Drawing.Imaging.ImageFormat.Emf;
                                break;

                            case "bmp":
                                Format = System.Drawing.Imaging.ImageFormat.Bmp;
                                break;

                            case "png":
                                Format = System.Drawing.Imaging.ImageFormat.Png;
                                break;

                            case "wmf":
                                Format = System.Drawing.Imaging.ImageFormat.Wmf;
                                break;

                            case "gif":
                                Format = System.Drawing.Imaging.ImageFormat.Gif;
                                break;
                        }
                        //
                       
                        img.Image.Save(path+"ResultedImage"+"." + ImageType, Format);
                    }
                    ifImageFound = false;
                }
            }
        }
    }
}
