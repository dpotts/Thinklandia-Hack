using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using Microsoft.Kinect;
using System.Runtime.InteropServices;
using System.Drawing.Imaging;
using System.Drawing;

namespace ThinklandiaHack
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        KinectSensor sensor;

        public MainWindow()
        {
            InitializeComponent();
            sensor = KinectSensor.KinectSensors[0];

            sensor.ColorStream.Enable(ColorImageFormat.RgbResolution640x480Fps30);
            sensor.DepthStream.Enable(DepthImageFormat.Resolution320x240Fps30);
            sensor.SkeletonStream.Enable();

            sensor.AllFramesReady += FramesReady;
            sensor.Start();
        }

        System.Drawing.Point GetJoint(JointType j, Skeleton S)
        {
            SkeletonPoint Sloc = S.Joints[j].Position;
            ColorImagePoint Cloc = sensor.MapSkeletonPointToColor(Sloc, ColorImageFormat.RgbResolution640x480Fps30);
            return new System.Drawing.Point(Cloc.X, Cloc.Y);
        }

        private void FramesReady(object sender, AllFramesReadyEventArgs e)
        {
            ColorImageFrame VFrame = e.OpenColorImageFrame();
            if (VFrame == null) return;
            byte[] pixelS = new byte[VFrame.PixelDataLength];
            Bitmap bmap = ImageToBitmap(VFrame);

            SkeletonFrame SFrame = e.OpenSkeletonFrame();
            if (SFrame == null) return;
            Skeleton[] Skeletons = new Skeleton[SFrame.SkeletonArrayLength];
            SFrame.CopySkeletonDataTo(Skeletons);

            foreach (Skeleton S in Skeletons)
            {
                if (S.TrackingState == SkeletonTrackingState.Tracked)
                {
                    SkeletonPoint Sloc = S.Joints[JointType.Head].Position;
                    ColorImagePoint Cloc = sensor.MapSkeletonPointToColor(Sloc, ColorImageFormat.RgbResolution640x480Fps30);
                    System.Drawing.Point p1 = GetJoint(JointType.Head, S);
                    System.Drawing.Point p2 = GetJoint(JointType.ShoulderCenter, S);
                    Graphics g = Graphics.FromImage(bmap);
                    g.DrawLine(Pens.Red, p1, p2);
                }
            }
        }

        private Bitmap ImageToBitmap(ColorImageFrame Image)
        {
            byte[] pixeldata = new byte[Image.PixelDataLength];
            Image.CopyPixelDataTo(pixeldata);
            Bitmap bmap = new Bitmap(Image.Width, Image.Height, System.Drawing.Imaging.PixelFormat.Format32bppRgb);
            BitmapData bmapdata = bmap.LockBits(new System.Drawing.Rectangle(0, 0, Image.Width, Image.Height), ImageLockMode.WriteOnly, bmap.PixelFormat);
            IntPtr ptr = bmapdata.Scan0;
            Marshal.Copy(pixeldata, 0, ptr, Image.PixelDataLength);
            bmap.UnlockBits(bmapdata);
            return bmap;
        }
    }


}
