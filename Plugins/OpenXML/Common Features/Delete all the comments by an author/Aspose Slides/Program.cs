// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Slides;
using System;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\..\Sample Files\";
            string FileName = FilePath + "Delete all the comments by an author.pptx";
            string author = "MZ";
            DeleteCommentsByAuthorInPresentation(FileName, author);
        }
        // Remove all the comments in the slides by a certain author.
        public static void DeleteCommentsByAuthorInPresentation(string fileName, string author)
        {
            if (String.IsNullOrEmpty(fileName) || String.IsNullOrEmpty(author))
                throw new ArgumentNullException("File name or author name is NULL!");
            //Instantiate a PresentationEx object that represents a PPTX file
            using (Presentation pres = new Presentation(fileName))
            {
              ICommentAuthor[] authors=  pres.CommentAuthors.FindByName(author);
              ICommentAuthor thisAuthor = authors[0];
              for (int i = thisAuthor.Comments.Count - 1; i >= 0;i-- )
              {
                  thisAuthor.Comments.RemoveAt(i);
              }
              pres.Save(fileName, Aspose.Slides.Export.SaveFormat.Pptx);  
            }

        }
    }
}
