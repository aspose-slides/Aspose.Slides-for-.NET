//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;
using System;

namespace AccessingSlideComments
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            Presentation pres = new Presentation(dataDir + "Comments.ppt");


            int i = 1;
            foreach (Slide slide in pres.Slides)
            {
                foreach (Comment comment in slide.SlideComments)
                {
                    Console.WriteLine("Slide :" + i.ToString() + " has comment: " + comment.Text + " with Author: " + comment.Author.Name + " posted on time :" + comment.CreatedTime + "\n");
                    i++;
                }
            }
            
            
        }
    }
}