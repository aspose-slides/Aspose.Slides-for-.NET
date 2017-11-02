using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/


namespace Aspose.Slides.Examples.CSharp.Presentations.Properties
{
    public class AddingEMZImagesToImageCollection
    {
        public static void Run()
        {
            //ExStart:AddingEMZImagesToImageCollection
           // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_PresentationProperties();
            Presentation p = new Presentation();
               ISlide s = p.Slides[0];
               // byte[] buffer=new byte();
              String imagePath=@"C:\Aspose Data\emf files\";
              byte[] data = GetCompressedData(imagePath + "2.emz");
             if (s != null)
        {
              if (s.Shapes != null)
          {
              IPPImage imgx = p.Images.AddImage(data);

              var m = s.Shapes.AddPictureFrame(ShapeType.Rectangle, 0, 0, p.SlideSize.Size.Width, p.SlideSize.Size.Height , imgx);
              p.Save("C:\\Asopse Data\\Saved.pptx", SaveFormat.Pptx);
          }
          }
         }
        

       //private byte[] GetCompressedData(string fileNameZip, byte[] buffer)
      private static byte[] GetCompressedData(string fileNameZip)
    {
        byte[] bufferZip = null;
      /*  byte[] buffer = null;

        FileStream f1 = new FileStream(fileName, FileMode.Open);
    byte[] buffer=f1.
        using (FileStream f = new FileStream(fileNameZip, FileMode.Create))
        {
            buffer = new byte[f.Length];
            using (var gz = new GZipStream(f, CompressionMode.Compress, false))
            {
                gz.Write(buffer, 0, buffer.Length);
            }
        }
    */
        using (FileStream f = new FileStream(fileNameZip, FileMode.Open))
        {
            bufferZip = new byte[f.Length];
            f.Read(bufferZip, 0, (int)f.Length);
        }

        return bufferZip;
        }
            //ExEnd:AddingEMZImagesToImageCollection
        }
    }
