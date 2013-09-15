using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;

using Microsoft.Kinect;

using System.Runtime.InteropServices;
using System.Drawing.Imaging;
using System.Media;
using System.Windows.Media.Imaging;
using System.Windows;
using System.Windows.Forms;

using Microsoft.Speech.AudioFormat;
using Microsoft.Speech.Recognition;
using System.IO;

//smoothing
//adding text to ecard?
//email?
//another picture?
//multiple people

namespace ANewHope
{
    public partial class Form1 : Form
    {
        KinectSensor sensor;

        int page = 1;
        //and the speech recognition engine (SRE)
        private SpeechRecognitionEngine speechRecognizer;
        //Get the speech recognizer (SR)
        private static RecognizerInfo GetKinectRecognizer()
        {
            Func<RecognizerInfo, bool> matchingFunc = r =>
            {
                string value;
                r.AdditionalInfo.TryGetValue("Kinect", out value);
                return "True".Equals(value, StringComparison.InvariantCultureIgnoreCase) && "en-US".Equals(r.Culture.Name, StringComparison.InvariantCultureIgnoreCase);
            };
            return SpeechRecognitionEngine.InstalledRecognizers().Where(matchingFunc).FirstOrDefault();
        }

        public Form1()
        {
            InitializeComponent();

            pictureBox2.Parent = pictureBox1;

            pictureBox3.Parent = pictureBox2;
            pictureBox4.Parent = pictureBox3; 

            foreach (var kinectSensor in KinectSensor.KinectSensors)
            {
                if (kinectSensor.Status == KinectStatus.Connected)
                {
                    sensor = kinectSensor;
                    break;
                }
            }
        }

        public void pictureBox5_Click(object sender, EventArgs e)
        {
            if (page == 1)
            {
                pictureBox6.Location = pictureBox2.Location;
                pictureBox1.Location = new Point(this.Location.X - (pictureBox1.Width + 140), pictureBox1.Location.Y);
                //pictureBox2.Location = new Point(this.Location.X - (pictureBox2.Width + 120), pictureBox2.Location.Y);
                page = 2;
            }
            else if (page == 2)
            {
                pictureBox1.Location = pictureBox6.Location;
                pictureBox6.Location = new Point(this.Width + 20, pictureBox6.Location.Y);
                page = 1;
            }
        }

        public void pictureBox7_Click(object sender, EventArgs e)
        {
            if (page == 1)
            {
                pictureBox6.Location = pictureBox2.Location;
                pictureBox1.Location = new Point(this.Location.X - (pictureBox1.Width + 140), pictureBox1.Location.Y);
                //pictureBox2.Location = new Point(this.Location.X - (pictureBox2.Width + 120), pictureBox2.Location.Y);
                page = 2;
            }
            else if (page == 2)
            {
                pictureBox1.Location = pictureBox6.Location;
                pictureBox6.Location = new Point(this.Width + 20, pictureBox6.Location.Y);
                page = 1;

            }

        }
        //Start streaming audio
        private void Start()
        {
            //set sensor audio source to variable
            var audioSource = sensor.AudioSource;
            //Set the beam angle mode - the direction the audio beam is pointing
            //we want it to be set to adaptive
            audioSource.BeamAngleMode = BeamAngleMode.Adaptive;
            //start the audiosource 
            var kinectStream = audioSource.Start();
            //configure incoming audio stream
            speechRecognizer.SetInputToAudioStream(
                kinectStream, new SpeechAudioFormatInfo(EncodingFormat.Pcm, 16000, 16, 1, 32000, 2, null));
            //make sure the recognizer does not stop after completing     
            speechRecognizer.RecognizeAsync(RecognizeMode.Multiple);
            //reduce background and ambient noise for better accuracy
            sensor.AudioSource.EchoCancellationMode = EchoCancellationMode.None;
            sensor.AudioSource.AutomaticGainControlEnabled = false;
        }

        //here is the fun part: create the speech recognizer
        private SpeechRecognitionEngine CreateSpeechRecognizer()
        {
            //set recognizer info
            RecognizerInfo ri = GetKinectRecognizer();
            //create instance of SRE
            SpeechRecognitionEngine sre;
            sre = new SpeechRecognitionEngine(ri.Id);

            //Now we need to add the words we want our program to recognise
            var grammar = new Choices();
            grammar.Add("next");
            grammar.Add("previous");

            grammar.Add("happy birthday");
            grammar.Add("capture");
            grammar.Add("congratulations");
            grammar.Add("rifflandia");

            grammar.Add("clear");







            //set culture - language, country/region
            var gb = new GrammarBuilder { Culture = ri.Culture };
            gb.Append(grammar);

            //set up the grammar builder
            var g = new Grammar(gb);
            sre.LoadGrammar(g);

            //Set events for recognizing, hypothesising and rejecting speech
            sre.SpeechRecognized += SreSpeechRecognized;
            sre.SpeechRecognitionRejected += SreSpeechRecognitionRejected;
            return sre;
        }

        //if speech is rejected
        private void RejectSpeech(RecognitionResult result)
        {
            Console.WriteLine("Rejected Speech");
        }

        private void SreSpeechRecognitionRejected(object sender, SpeechRecognitionRejectedEventArgs e)
        {
            RejectSpeech(e.Result);
        }


        //Speech is recognised
        private void SreSpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            //Very important! - change this value to adjust accuracy - the higher the value
            //the more accurate it will have to be, lower it if it is not recognizing you


            if (e.Result.Confidence < .80)
            {
                RejectSpeech(e.Result);
            }
            else
            {
                //and finally, here we set what we want to happen when 
                //the SRE recognizes a word
                switch (e.Result.Text.ToUpperInvariant())
                {
                    case "NEXT":
                        this.pictureBox5_Click(null, null);
                        break;
                    case "PREVIOUS":
                        this.pictureBox7_Click(null, null);
                        break;
                    case "CAPTURE":
                        break;
                    case "HAPPY BIRTHDAY":
                         textBox1.Text = "HAPPY BIRTHDAY!!!";
                          break;
                      case "CONGRATULATIONS":
                          textBox1.Text = "CONGRATULATIONS!!!";
                          break;
                      case "RIFFLANDIA":
                          textBox1.Text = "ENJOY RIFFLANDIA!";
                          break;
                      case "CLEAR":
                          textBox1.Text = null;
                          break;
                    default:
                        break;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (sensor != null)
            {
                sensor.ColorStream.Enable(ColorImageFormat.RgbResolution640x480Fps30);
                sensor.DepthStream.Enable(DepthImageFormat.Resolution320x240Fps30);
                sensor.SkeletonStream.EnableTrackingInNearRange = false;
                sensor.SkeletonStream.TrackingMode = SkeletonTrackingMode.Seated;
                sensor.SkeletonStream.Enable();
                sensor.AllFramesReady += FramesReady;
                sensor.Start();
                speechRecognizer = CreateSpeechRecognizer();
                Start();
            }
        }

        private Bitmap ExtractBodyPartBitmap(KinectSensor sensor, Skeleton skeleton, Bitmap bmap, JointType jointType, float widthMeters, float heightMeters)
        {
            ColorImageStream stream = sensor.ColorStream;
            ColorImageFormat imageFormat = stream.Format;
            int stride = stream.FrameWidth * stream.FrameBytesPerPixel;
            CoordinateMapper map = new CoordinateMapper(sensor);

            // Get the position of the joint
            SkeletonPoint skeletonPoint = skeleton.Joints[jointType].Position;

            // Map joint position to pixel coordinates
            ColorImagePoint colorPoint = map.MapSkeletonPointToColorPoint(skeletonPoint, imageFormat);

            // Estimate a cutout rectangle centered on the pixel coordinates
            //Int32Rect rect = EstimateCutoutRect(stream, colorPoint, skeletonPoint.Z, widthMeters, heightMeters);
            Bitmap bitmap = DrawHead(stream, bmap, colorPoint, skeletonPoint.Z, widthMeters, heightMeters);

            // Cut out the appropriate part of the image
            //WriteableBitmap bitmap = new WriteableBitmap(rect.Width, rect.Height, 96, 96, PixelFormats.Bgr32, null);
            //bitmap.WritePixels(rect, imageData, stride, 0, 0);
            return bitmap;
        }

        private Bitmap DrawHead(ColorImageStream stream, Bitmap bmap, ColorImagePoint colorPoint, float depthMeters, float widthMeters, float heightMeters)
        {
            float focalLength = stream.NominalFocalLengthInPixels;
            int widthPixels = (int)((widthMeters * focalLength) / depthMeters);
            int heightPixels = (int)((heightMeters * focalLength) / depthMeters);

            int x = System.Math.Max(0, colorPoint.X - (widthPixels / 2));
            int y = System.Math.Max(0, colorPoint.Y - (heightPixels / 2));
            widthPixels = System.Math.Min(widthPixels, stream.FrameWidth - x);
            heightPixels = System.Math.Min(heightPixels, stream.FrameHeight - y);

            Pen skyBluePen = new Pen(Brushes.White);

            // Set the pen's width.
            skyBluePen.Width = 5.0F;

            // Set the LineJoin property.
            skyBluePen.LineJoin = System.Drawing.Drawing2D.LineJoin.Bevel;

            Bitmap bitmap = new Bitmap(widthPixels, heightPixels, PixelFormat.Format32bppRgb);

            //Graphics g = Graphics.FromImage(bmap);
            //g.DrawEllipse(skyBluePen, x , (y + 15.0f), widthPixels, heightPixels);
            cropImageToCircle(bmap, x, (y + 15.0f), widthPixels, heightPixels);

            return bitmap;
        }

        private void cropImageToCircle(Bitmap bmap, float circleStartX, float circleStartY, float width, float height)
        {
            Rectangle aRect = new Rectangle((int)circleStartX, (int)circleStartY, (int)width, (int)height);

            float wScale = width;
            float hScale = height;
            if (page == 1)
            {
                wScale = (float)110.0;
                hScale = (float)150.0;
            }
            else if (page == 2)
            {
                wScale = (float)170.0;
                hScale = (float)230.0;
            }

            // Check if aRect is within bmap bounds before using Clone() method 
            if ((bmap.Width >= aRect.Right) && (bmap.Height >= aRect.Bottom))
            {
                Bitmap cropped = bmap.Clone(aRect, bmap.PixelFormat);
                TextureBrush tb = new TextureBrush(cropped);
                float xScale = (float)(wScale / width);
                tb.ScaleTransform(xScale, xScale);

                Graphics g = Graphics.FromImage(bmap);
                Color c = Color.Red;
                g.Clear(c);
                g.FillEllipse(tb, 0, 0, wScale, hScale);
            }
            //Bitmap final = new Bitmap((int)width, (int)height);
            //Graphics g = Graphics.FromImage(final);
            //g.FillEllipse(tb, 0, 0, width, height);
        }

        void FramesReady(object sender, AllFramesReadyEventArgs e)
        {
            ColorImageFrame VFrame = e.OpenColorImageFrame();
            if (VFrame == null) return;
            byte[] pixelS = new byte[VFrame.PixelDataLength];
            Bitmap bmap = ImageToBitmap(VFrame);


            SkeletonFrame SFrame = e.OpenSkeletonFrame();
            if (SFrame == null) return;

            Graphics g = Graphics.FromImage(bmap);
            Skeleton[] Skeletons = new Skeleton[SFrame.SkeletonArrayLength];
            SFrame.CopySkeletonDataTo(Skeletons);

            foreach (Skeleton S in Skeletons)
            {
                if (S.TrackingState == SkeletonTrackingState.Tracked)
                {

                    ExtractBodyPartBitmap(this.sensor, S, bmap, JointType.Head, 0.17f, 0.26f);

                }

            }

            // Release resources used by VFrame and SFrame objects 
            VFrame.Dispose();
            SFrame.Dispose();

            pictureBox3.Image = bmap;
            bmap.MakeTransparent(Color.Red);
            Size tempSize = bmap.Size;

            if (page == 1)
            {
                pictureBox2.Parent = pictureBox1;

                pictureBox3.Parent = pictureBox2;
                pictureBox4.Parent = pictureBox3;

                pictureBox3.Location = new Point(pictureBox2.Width / 2 - 15, 135);

                tempSize.Width = 110;
                tempSize.Height = 150;
            }
            else if (page == 2)
            {
                pictureBox3.Parent = pictureBox6;

                pictureBox3.Location = new Point(pictureBox2.Width / 2 - 120, 80);

                tempSize.Width = 170;
                tempSize.Height = 230;
            }

            pictureBox3.SizeMode = PictureBoxSizeMode.Normal;
            pictureBox3.Size = tempSize;

            pictureBox2.BackColor = Color.Transparent;

            pictureBox1.BackColor = Color.Transparent;
            pictureBox4.Location = new Point(-300, -300);
            pictureBox4.Size = new Size(130, 175);

        }

        void DrawBone(JointType j1, JointType j2, Skeleton S, Graphics g)
        {
            Point p1 = GetJoint(j1, S);
            Point p2 = GetJoint(j2, S);
            g.DrawLine(Pens.Red, p1, p2);
        }

        Point GetJoint(JointType j, Skeleton S)
        {
            SkeletonPoint Sloc = S.Joints[j].Position;
            ColorImagePoint Cloc = sensor.CoordinateMapper.MapSkeletonPointToColorPoint(Sloc, ColorImageFormat.RawBayerResolution640x480Fps30);

            return new Point(Cloc.X, Cloc.Y);
        }



        Bitmap ImageToBitmap(ColorImageFrame Image)
        {
            byte[] pixelData = new byte[Image.PixelDataLength];
            Image.CopyPixelDataTo(pixelData);
            Bitmap bmap = new Bitmap(
                   Image.Width,
                   Image.Height,
                   PixelFormat.Format32bppRgb);
            BitmapData bmapS = bmap.LockBits(
              new Rectangle(0, 0,
                         Image.Width, Image.Height),
              ImageLockMode.WriteOnly,
              bmap.PixelFormat);
            IntPtr ptr = bmapS.Scan0;
            Marshal.Copy(pixelData, 0, ptr,
                       Image.PixelDataLength);
            bmap.UnlockBits(bmapS);
            return bmap;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            sensor.Stop();
        }


    }
}
