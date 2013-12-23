//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;

namespace AccessExistingTable
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Instantiate a Presentation object that represents a PPT file
            Presentation pres = new Presentation(dataDir + "demo.ppt");


            //Accessing a slide using its slide position
            Slide slide = pres.GetSlideByPosition(2);


            //Setting table object to null
            Table table = null;


            //Iterating through all shapes unless the desired table is found
            for (int i = 0; i < slide.Shapes.Count; i++)
            {
                if (slide.Shapes[i] is Table)
                {
                    table = (Table)slide.Shapes[i];


                    if (table.AlternativeText.Equals("myTable"))
                    {
                        System.Console.WriteLine("Table Found");
                        break;
                    }
                }
            }


            //Adding a new row in the table
            table.AddRow();


            //Writing the presentation as a PPT file
            pres.Write(dataDir + "modified.ppt");

        }
    }
}