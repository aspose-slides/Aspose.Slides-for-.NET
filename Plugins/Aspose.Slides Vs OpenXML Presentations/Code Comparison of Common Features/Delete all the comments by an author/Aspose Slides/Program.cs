// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using System;
using Aspose.Slides.Pptx;

namespace Aspose_Slides
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName = "Delete all the comments by an author.pptx";
            string author = "MZ";
            DeleteCommentsByAuthorInPresentation(fileName, author);
        }
        // Remove all the comments in the slides by a certain author.
        public static void DeleteCommentsByAuthorInPresentation(string fileName, string author)
        {
            if (String.IsNullOrEmpty(fileName) || String.IsNullOrEmpty(author))
                throw new ArgumentNullException("File name or author name is NULL!");
            //Instantiate a PresentationEx object that represents a PPTX file
            using (PresentationEx pres = new PresentationEx(fileName))
            {
              CommentAuthorEx[] authors=  pres.CommentAuthors.FindByName(author);
              CommentAuthorEx thisAuthor = authors[0];
              for (int i = thisAuthor.Comments.Count - 1; i >= 0;i-- )
              {
                  thisAuthor.Comments.RemoveAt(i);
              }
              pres.Save(fileName, Aspose.Slides.Export.SaveFormat.Pptx);  
            }

        }
    }
}
