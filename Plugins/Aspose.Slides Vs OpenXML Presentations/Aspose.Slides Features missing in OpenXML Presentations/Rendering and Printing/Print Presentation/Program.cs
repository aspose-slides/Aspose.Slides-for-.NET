// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Slides.Pptx;

namespace Print_Presentation
{
    class Program
    {
        static void Main(string[] args)
        {
            PrintByDefaultPrinter();
            PrintBySpecificPrinter();
        }
        public static void PrintByDefaultPrinter()
        {
            string MyDir = @"Files\";
            //Load the presentation
            PresentationEx asposePresentation = new PresentationEx(MyDir + "Print.pptx");

            //Call the print method to print whole presentation to the default printer
            asposePresentation.Print();

        }
        public static void PrintBySpecificPrinter()
        {
            string MyDir = @"Files\";
            //Load the presentation
            PresentationEx asposePresentation = new PresentationEx(MyDir + "Print.pptx");

            //Call the print method to print whole presentation to the desired printer
            asposePresentation.Print("LaserJet1100");

        }
    }
}
