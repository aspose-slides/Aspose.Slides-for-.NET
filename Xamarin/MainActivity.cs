using System;
using System.IO;
using Android.App;
using Android.OS;
using Android.Support.V7.App;
using Android.Widget;

namespace PresentationViewer
{
    [Activity(Label = "@string/app_name", Theme = "@style/AppTheme.NoActionBar", MainLauncher = true)]
    public class MainActivity : AppCompatActivity
    {
        private Button buttonNext;
        private Button buttonPrev;
        ImageView imageView;
        private Aspose.Slides.Presentation presentation;
        private int currentSlideNumber;

        protected override void OnCreate(Bundle savedInstanceState)
        {
            base.OnCreate(savedInstanceState);
            SetContentView(Resource.Layout.activity_main);
        }
        
        protected override void OnResume()
        {
            base.OnResume();

            LoadPresentation();
            currentSlideNumber = 0;

            if (buttonNext == null)
            {
                buttonNext = FindViewById<Button>(Resource.Id.buttonNext);
            }

            if (buttonPrev == null)
            {
                buttonPrev = FindViewById<Button>(Resource.Id.buttonPrev);
            }

            if(imageView == null)
            {
                imageView= FindViewById<ImageView>(Resource.Id.imageView);
            }

            buttonNext.Click += ButtonNext_Click;
            buttonPrev.Click += ButtonPrev_Click;
            imageView.Touch += ImageView_Touch;

            RefreshButtonsStatus();
            ShowSlide(currentSlideNumber);
        }

        private void ButtonNext_Click(object sender, System.EventArgs e)
        {
            if (currentSlideNumber > (presentation.Slides.Count - 1))
            {
                return;
            }

            ShowSlide(++currentSlideNumber);
            RefreshButtonsStatus();
        }

        private void ButtonPrev_Click(object sender, System.EventArgs e)
        {
            if (currentSlideNumber == 0)
            {
                return;
            }

            ShowSlide(--currentSlideNumber);
            RefreshButtonsStatus();
        }

        private void ImageView_Touch(object sender, Android.Views.View.TouchEventArgs e)
        {
            int[] location = new int[2];
            imageView.GetLocationOnScreen(location);

            int x = (int)e.Event.GetX();
            int y = (int)e.Event.GetY();
            int posX = x - location[0];
            int posY = y - location[0];

            Aspose.Slides.Drawing.Xamarin.Size presSize = presentation.SlideSize.Size.ToSize();
            float coeffX = (float)presSize.Width / imageView.Width;
            float coeffY = (float)presSize.Height / imageView.Height;

            int presPosX = (int)(posX * coeffX);
            int presPosY = (int)(posY * coeffY);

            int width = presSize.Width / 50;
            int height = width;

            Aspose.Slides.IAutoShape ellipse = presentation.Slides[currentSlideNumber].Shapes.AddAutoShape(Aspose.Slides.ShapeType.Ellipse, presPosX, presPosY, width, height);
            ellipse.FillFormat.FillType = Aspose.Slides.FillType.Solid;
            
            Random random = new Random();
            Aspose.Slides.Drawing.Xamarin.Color slidesColor = Aspose.Slides.Drawing.Xamarin.Color.FromArgb(random.Next(256), random.Next(256), random.Next(256));
            ellipse.FillFormat.SolidFillColor.Color = slidesColor;

            ShowSlide(currentSlideNumber);
        }

        protected override void OnPause()
        {
            base.OnPause();

            if (buttonNext != null)
            {
                buttonNext.Dispose();
                buttonNext = null;
            }

            if (buttonPrev != null)
            {
                buttonPrev.Dispose();
                buttonPrev = null;
            }

            if(imageView != null)
            {
                imageView.Dispose();
                imageView = null;
            }

            DisposePresentation();
        }

        private void RefreshButtonsStatus()
        {
            buttonNext.Enabled = currentSlideNumber < (presentation.Slides.Count - 1);
            buttonPrev.Enabled = currentSlideNumber > 0;
        }

        private void ShowSlide(int slideNumber)
        {
            Aspose.Slides.Drawing.Xamarin.Size size = presentation.SlideSize.Size.ToSize();
            Aspose.Slides.Drawing.Xamarin.Bitmap bitmap = presentation.Slides[slideNumber].GetThumbnail(size);
            imageView.SetImageBitmap(bitmap.ToNativeBitmap());
        }

        private void LoadPresentation()
        {
            if(presentation != null)
            {
                return;
            }

            using (Stream input = Assets.Open("HelloWorld.pptx"))
            {
                presentation = new Aspose.Slides.Presentation(input);
            }
        }

        private void DisposePresentation()
        {
            if(presentation == null)
            {
                return;
            }

            presentation.Dispose();
            presentation = null;
        }
    }
}

