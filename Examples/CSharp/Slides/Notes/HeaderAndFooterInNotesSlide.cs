using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Slides.Notes
{
	class HeaderAndFooterInNotesSlide
	{
     public static void Run()
		{
			//ExStart:HeaderAndFooterInNotesSlide
			string dataDir = RunExamples.GetDataDir_Slides_Presentations_Notes();
			using (Presentation presentation = new Presentation(dataDir + "presentation.pptx"))
			{
				// Change Header and Footer settings for notes master and all notes slides
				IMasterNotesSlide masterNotesSlide = presentation.MasterNotesSlideManager.MasterNotesSlide;
				if (masterNotesSlide != null)
				{
					IMasterNotesSlideHeaderFooterManager headerFooterManager = masterNotesSlide.HeaderFooterManager;

					headerFooterManager.SetHeaderAndChildHeadersVisibility(true); // make the master notes slide and all child Footer placeholders visible
					headerFooterManager.SetFooterAndChildFootersVisibility(true); // make the master notes slide and all child Header placeholders visible
					headerFooterManager.SetSlideNumberAndChildSlideNumbersVisibility(true); // make the master notes slide and all child SlideNumber placeholders visible
					headerFooterManager.SetDateTimeAndChildDateTimesVisibility(true); // make the master notes slide and all child Date and time placeholders visible

					headerFooterManager.SetHeaderAndChildHeadersText("Header text"); // set text to master notes slide and all child Header placeholders
					headerFooterManager.SetFooterAndChildFootersText("Footer text"); // set text to master notes slide and all child Footer placeholders
					headerFooterManager.SetDateTimeAndChildDateTimesText("Date and time text"); // set text to master notes slide and all child Date and time placeholders
				}

				// Change Header and Footer settings for first notes slide only
				INotesSlide notesSlide = presentation.Slides[0].NotesSlideManager.NotesSlide;
				if (notesSlide != null)
				{
					INotesSlideHeaderFooterManager headerFooterManager = notesSlide.HeaderFooterManager;
					if (!headerFooterManager.IsHeaderVisible)
						headerFooterManager.SetHeaderVisibility(true); // make this notes slide Header placeholder visible

					if (!headerFooterManager.IsFooterVisible)
						headerFooterManager.SetFooterVisibility(true); // make this notes slide Footer placeholder visible

					if (!headerFooterManager.IsSlideNumberVisible)
						headerFooterManager.SetSlideNumberVisibility(true); // make this notes slide SlideNumber placeholder visible

					if (!headerFooterManager.IsDateTimeVisible)
						headerFooterManager.SetDateTimeVisibility(true); // make this notes slide Date-time placeholder visible

					headerFooterManager.SetHeaderText("New header text"); // set text to notes slide Header placeholder
					headerFooterManager.SetFooterText("New footer text"); // set text to notes slide Footer placeholder
					headerFooterManager.SetDateTimeText("New date and time text"); // set text to notes slide Date-time placeholder
				}
				presentation.Save(dataDir + "testresult.pptx",SaveFormat.Pptx);
			}
		
		  }
		
		//ExEnd:HeaderAndFooterInNotesSlide
	}
}