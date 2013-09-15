﻿using System;
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

namespace ANewHope
{
    public partial class Form1 : Form
    {
        //Leprechaun: 820x, 650y, 210w, 350h

        KinectSensor sensor;
        public Form1()
        {
            InitializeComponent();
            foreach (var kinectSensor in KinectSensor.KinectSensors)
            {
                if (kinectSensor.Status == KinectStatus.Connected)
                {
                    sensor = kinectSensor;
                    pictureBox2.Parent = pictureBox1;

                    pictureBox3.Parent = pictureBox2;
                    pictureBox4.Parent = pictureBox3;
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
            }
       
// I dunno if we'll need this later but maybe. lets just keep it in because its going the exception crap that we may need
    //        if (sensor == null)
    //        {
  /*              sensor.Start();
                try
                {
                    sensor.ColorStream.Enable(ColorImageFormat.RgbResolution640x480Fps30);
                    sensor.DepthStream.Enable(DepthImageFormat.Resolution320x240Fps30);
                    try
                    {
                        sensor.DepthStream.Range = DepthRange.Near;
                        sensor.SkeletonStream.EnableTrackingInNearRange = true;
                    }
                    catch (InvalidOperationException)
                    {
                        sensor.DepthStream.Range = DepthRange.Default;
                        sensor.SkeletonStream.EnableTrackingInNearRange = false;
                    }
                    sensor.SkeletonStream.TrackingMode = SkeletonTrackingMode.Seated;
                    sensor.SkeletonStream.Enable();
                    sensor.AllFramesReady += FramesReady;

 
                }
                catch (InvalidOperationException)
                {
                }*/
    //        }
        }

        private static Bitmap ExtractBodyPartBitmap(KinectSensor sensor,Skeleton skeleton, Bitmap bmap, JointType jointType,float widthMeters,float heightMeters)
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

        private static Bitmap DrawHead(ColorImageStream stream, Bitmap bmap, ColorImagePoint colorPoint, float depthMeters, float widthMeters, float heightMeters)
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

        private static void cropImageToCircle(Bitmap bmap, float circleStartX, float circleStartY, float width, float height)
        {
            Rectangle aRect = new Rectangle((int)circleStartX, (int)circleStartY, (int)width, (int)height);

            // Check if aRect is within bmap bounds before using Clone() method 
            if ((bmap.Width >= aRect.Right) && (bmap.Height >= aRect.Bottom))
            {
                Bitmap cropped = bmap.Clone(aRect, bmap.PixelFormat);
                TextureBrush tb = new TextureBrush(cropped);

                
                Graphics g = Graphics.FromImage(bmap);
                Color c = Color.Red;
                g.Clear(c);
                g.FillEllipse(tb, 0, 0, width, height);
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
            
            tempSize.Width = 130;
            tempSize.Height = 175;

            pictureBox3.SizeMode = PictureBoxSizeMode.Normal;
            pictureBox3.Size = tempSize;
            pictureBox3.Location = new Point(pictureBox2.Width / 2-20, 120);
            pictureBox2.BackColor = Color.Transparent;

            pictureBox1.BackColor = Color.Transparent;
            pictureBox4.Location = new Point(0, 0);
            pictureBox4.Size = new Size(120, 175);
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
