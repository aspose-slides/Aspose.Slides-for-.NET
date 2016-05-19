// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\..\Sample Files\";
            string FileName = FilePath + "Delete all the comments by an author.pptx";
            string author = "Zeeshan Shafqat";
            DeleteCommentsByAuthorInPresentation(FileName, author);
        }
        // Remove all the comments in the slides by a certain author.
        public static void DeleteCommentsByAuthorInPresentation(string fileName, string author)
        {
            if (String.IsNullOrEmpty(fileName) || String.IsNullOrEmpty(author))
                throw new ArgumentNullException("File name or author name is NULL!");

            using (PresentationDocument doc = PresentationDocument.Open(fileName, true))
            {
                // Get the specifed comment author.
                IEnumerable<CommentAuthor> commentAuthors =
                    doc.PresentationPart.CommentAuthorsPart.CommentAuthorList.Elements<CommentAuthor>()
                    .Where(e => e.Name.Value.Equals(author));

                // Iterate through all the matching authors.
                foreach (CommentAuthor commentAuthor in commentAuthors)
                {
                    UInt32Value authorId = commentAuthor.Id;

                    // Iterate through all the slides and get the slide parts.
                    foreach (SlidePart slide in doc.PresentationPart.SlideParts)
                    {
                        SlideCommentsPart slideCommentsPart = slide.SlideCommentsPart;
                        // Get the list of comments.
                        if (slideCommentsPart != null && slide.SlideCommentsPart.CommentList != null)
                        {
                            IEnumerable<Comment> commentList =
                                slideCommentsPart.CommentList.Elements<Comment>().Where(e => e.AuthorId == authorId.Value);
                            List<Comment> comments = new List<Comment>();
                            comments = commentList.ToList<Comment>();

                            foreach (Comment comm in comments)
                            {
                                // Delete all the comments by the specified author.

                                slideCommentsPart.CommentList.RemoveChild<Comment>(comm);
                            }

                            // If the commentPart has no existing comment.
                            if (slideCommentsPart.CommentList.ChildElements.Count == 0)
                                // Delete this part.
                                slide.DeletePart(slideCommentsPart);
                        }
                    }
                    // Delete the comment author from the comment authors part.
                    doc.PresentationPart.CommentAuthorsPart.CommentAuthorList.RemoveChild<CommentAuthor>(commentAuthor);
                }
            }
        }

    }
}
