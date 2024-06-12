using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;

/*
The following code sample demonstrates how to handle text replacing results using the IFindResultCallback interface:
*/

namespace CSharp.Text
{
    class FindAndReplaceTextResult
    {
        public static void Run()
        {
            string presentationName = Path.Combine(RunExamples.GetDataDir_Text(), "TextReplaceExample.pptx");
            string outPath = Path.Combine(RunExamples.OutPath, "TextReplaceExampleReplace-out.pptx");

            using (Presentation pres = new Presentation(presentationName))
            {
                // Create callback.
                FindResultCallback callback = new FindResultCallback();

                // Replace all words "[this block] ". 
                pres.ReplaceText("[this block] ", "my text", new TextSearchOptions(), callback);

                // Output the number of found fragments of the given text. 
                Console.WriteLine("number of found fragments = {0}", callback.Count);

                // Output data for each word "[this block] " found. 
                foreach (WordInfo info in callback.Words)
                {
                    Console.WriteLine("Text = {0}, Position = {1}", info.FoundText, info.TextPosition);
                }

                pres.Save(outPath, SaveFormat.Pptx);
            }
        }

        /// <summary>
        /// Class that provides information about all found occurrences of a given text.
        /// </summary>
        internal class FindResultCallback : IFindResultCallback
        {
            // Array of retrieved text information.
            public readonly List<WordInfo> Words = new List<WordInfo>();

            /// <summary>
            /// The number of matches found to a given text.
            /// </summary>
            public int Count
            {
                get { return Words.Count; }
            }

            /// <summary>
            /// Gets all slides in which the given text was found.
            /// </summary>
            public int[] SlideNumbers
            {
                get
                {
                    var slideNumbers = new List<int>();
                    foreach (var element in Words)
                    {
                        int slideNumber = ((Slide)element.TextFrame.Slide).SlideNumber;
                        if (!slideNumbers.Contains(slideNumber))
                            slideNumbers.Add(slideNumber);
                    }
                    return slideNumbers.ToArray();
                }
            }

            /// <summary>
            /// Gets all occurrences of the found text on the slide.
            /// </summary>
            /// <param name="slideNumber">Slide number</param>
            public WordInfo[] GetElemensForSlide(int slideNumber)
            {
                var foundElements = new List<WordInfo>();
                foreach (var element in Words)
                {
                    if (((Slide)element.TextFrame.Slide).SlideNumber == slideNumber)
                        foundElements.Add(element);
                }
                return foundElements.ToArray();
            }

            /// <summary>
            /// Callback method that receives data about the found text.
            /// </summary>
            /// <param name="textFrame">The <see cref="ITextFrame"/> in which the text was found.</param>
            /// <param name="sourceText">The source text in which the text was found.</param>
            /// <param name="foundText">The found text.</param>
            /// <param name="textPosition">The position of the found text.</param>
            public void FoundResult(ITextFrame textFrame, string oldText, string foundText, int textPosition)
            {
                Words.Add(new WordInfo(textFrame, oldText, foundText, textPosition));
            }
        }

        /// <summary>
        /// Class providing information about each text found in the presentation.
        /// </summary>
        public class WordInfo
        {
            internal WordInfo(ITextFrame textFrame, string sourceText, string foundText, int textPosition)
            {
                TextFrame = textFrame;
                SourceText = sourceText;
                FoundText = foundText;
                TextPosition = textPosition;
            }

            /// <summary>
            /// Gets found text.
            /// </summary>
            public string FoundText { get; }

            /// <summary>
            /// Gets the source text for the TextFrame in which the text was found.
            /// </summary>
            public string SourceText { get; }

            /// <summary>
            /// The position of the found text in the text frame.
            /// </summary>
            public int TextPosition { get; }

            /// <summary>
            /// The text frame in which the text was found.
            /// </summary>
            public ITextFrame TextFrame { get; }
        }
    }
}
