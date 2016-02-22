using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides.Pptx;
using Aspose.Slides;

namespace Protect_Presentation
{
    class Program
    {
        static void Main(string[] args)
        {
            ApplyingProtection();
            RemovingProtection();
        }
        static void ApplyingProtection()
        {
            string MyDir = @"Files\";
            //Instatiate Presentation class that represents a PPTX file
            PresentationEx pTemplate = new PresentationEx(MyDir + "Applying Protection.pptx");//Instatiate Presentation class that represents a PPTX file
           

            //ISlide object for accessing the slides in the presentation
            SlideEx slide = pTemplate.Slides[0];

            //IShape object for holding temporary shapes
            ShapeEx shape;

            //Traversing through all the slides in the presentation
            for (int slideCount = 0; slideCount < pTemplate.Slides.Count; slideCount++)
            {
                slide = pTemplate.Slides[slideCount];

                //Travesing through all the shapes in the slides
                for (int count = 0; count < slide.Shapes.Count; count++)
                {
                    shape = slide.Shapes[count];

                    //if shape is autoshape
                    if (shape is AutoShapeEx)
                    {
                        //Type casting to Auto shape and  getting auto shape lock
                        AutoShapeEx Ashp = shape as AutoShapeEx;
                        AutoShapeLockEx AutoShapeLock = Ashp.ShapeLock;

                        //Applying shapes locks
                        AutoShapeLock.PositionLocked = true;
                        AutoShapeLock.SelectLocked = true;
                        AutoShapeLock.SizeLocked = true;
                    }

                    //if shape is group shape
                    else if (shape is GroupShapeEx)
                    {
                        //Type casting to group shape and  getting group shape lock
                        GroupShapeEx Group = shape as GroupShapeEx;
                        GroupShapeLockEx groupShapeLock = Group.ShapeLock;

                        //Applying shapes locks
                        groupShapeLock.GroupingLocked = true;
                        groupShapeLock.PositionLocked = true;
                        groupShapeLock.SelectLocked = true;
                        groupShapeLock.SizeLocked = true;
                    }

                    //if shape is a connector
                    else if (shape is ConnectorEx)
                    {
                        //Type casting to connector shape and  getting connector shape lock
                        ConnectorEx Conn = shape as ConnectorEx;
                        ConnectorLockEx ConnLock = Conn.ShapeLock;

                        //Applying shapes locks
                        ConnLock.PositionMove = true;
                        ConnLock.SelectLocked = true;
                        ConnLock.SizeLocked = true;
                    }

                    //if shape is picture frame
                    else if (shape is PictureFrameEx)
                    {
                        //Type casting to picture frame shape and  getting picture frame shape lock
                        PictureFrameEx Pic = shape as PictureFrameEx;
                        PictureFrameLockEx PicLock = Pic.ShapeLock;

                        //Applying shapes locks
                        PicLock.PositionLocked = true;
                        PicLock.SelectLocked = true;
                        PicLock.SizeLocked = true;
                    }
                }
            }
            //Saving the presentation file
            pTemplate.Save(MyDir+"ProtectedSample.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
        }
        static void RemovingProtection()
        {
            string MyDir = @"Files\";
            //Open the desired presentation
            PresentationEx pTemplate = new PresentationEx(MyDir + "ProtectedSample.pptx");

            //ISlide object for accessing the slides in the presentation
            SlideEx slide = pTemplate.Slides[0];

            //IShape object for holding temporary shapes
            ShapeEx shape;

            //Traversing through all the slides in presentation
            for (int slideCount = 0; slideCount < pTemplate.Slides.Count; slideCount++)
            {
                slide = pTemplate.Slides[slideCount];

                //Travesing through all the shapes in the slides
                for (int count = 0; count < slide.Shapes.Count; count++)
                {
                    shape = slide.Shapes[count];

                    //if shape is autoshape
                    if (shape is AutoShapeEx)
                    {
                        //Type casting to Auto shape and  getting auto shape lock
                        AutoShapeEx Ashp = shape as AutoShapeEx;
                        AutoShapeLockEx AutoShapeLock = Ashp.ShapeLock;

                        //Applying shapes locks
                        AutoShapeLock.PositionLocked = false;
                        AutoShapeLock.SelectLocked = false;
                        AutoShapeLock.SizeLocked = false;
                    }

                    //if shape is group shape
                    else if (shape is GroupShapeEx)
                    {
                        //Type casting to group shape and  getting group shape lock
                        GroupShapeEx Group = shape as GroupShapeEx;
                        GroupShapeLockEx groupShapeLock = Group.ShapeLock;

                        //Applying shapes locks
                        groupShapeLock.GroupingLocked = false;
                        groupShapeLock.PositionLocked = false;
                        groupShapeLock.SelectLocked = false;
                        groupShapeLock.SizeLocked = false;
                    }

                    //if shape is Connector shape
                    else if (shape is ConnectorEx)
                    {
                        //Type casting to connector shape and  getting connector shape lock
                        ConnectorEx Conn = shape as ConnectorEx;
                        ConnectorLockEx ConnLock = Conn.ShapeLock;

                        //Applying shapes locks
                        ConnLock.PositionMove = false;
                        ConnLock.SelectLocked = false;
                        ConnLock.SizeLocked = false;
                    }

                    //if shape is picture frame
                    else if (shape is PictureFrameEx)
                    {
                        //Type casting to pitcture frame shape and  getting picture frame shape lock
                        PictureFrameEx Pic = shape as PictureFrameEx;
                        PictureFrameLockEx PicLock = Pic.ShapeLock;

                        //Applying shapes locks
                        PicLock.PositionLocked = false;
                        PicLock.SelectLocked = false;
                        PicLock.SizeLocked = false;
                    }
                }

            }
            //Saving the presentation file
            pTemplate.Save(MyDir+"RemoveProtectionSample.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
        }
    }
}
