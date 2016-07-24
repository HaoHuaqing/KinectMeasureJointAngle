using Microsoft.Kinect;
using Oraycn.MCapture;
using Oraycn.MFile;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace KinectCoordinateMapping
{

    public partial class MainWindow : Window
    {
        //定义kinect初始变量
        KinectSensor _sensor;
        MultiSourceFrameReader _reader;
        IList<Body> _bodies;
        CameraMode _mode = CameraMode.Color;//定义kinect初始模式彩色图

        //定义关节，标志位
        private bool shoulderrightflag = false;
        private bool elbowrightflag = false;
        private bool shoulderleftflag = false;
        private bool elbowleftflag = false;
        private WriteableBitmap colorBitmap = null;
        System.Windows.Point ShoulderRight = new System.Windows.Point();
        System.Windows.Point ElbowRight = new System.Windows.Point();
        System.Windows.Point WristRight = new System.Windows.Point();
        System.Windows.Point SpineShoulder = new System.Windows.Point();
        System.Windows.Point ShoulderLeft = new System.Windows.Point();
        System.Windows.Point ElbowLeft = new System.Windows.Point();
        System.Windows.Point WristLeft = new System.Windows.Point();
        double rightshoulderangle;
        double rightelbowangle;

        //用于录制
        private SilenceVideoFileMaker silenceVideoFileMaker;
        private IDesktopCapturer desktopCapturer;
        private int frameRate = 10; // 采集视频的帧频
        private bool sizeRevised = false;// 是否需要将图像帧的长宽裁剪为4的整数倍
        private volatile bool isRecording = false;//volatile作为指令关键字，确保本条指令不会因编译器的优化而省略，且要求每次直接读值.

        //用于获取应用程序矩形
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetWindowRect(HandleRef hWnd, out RECT lpRect);
        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
        public MainWindow()
        {
            InitializeComponent();
            //视频录制dll加载
            Oraycn.MCapture.GlobalUtil.SetAuthorizedUser("FreeUser", "");
            Oraycn.MFile.GlobalUtil.SetAuthorizedUser("FreeUser", "");
            this.colorBitmap = new WriteableBitmap(1920, 1080, 96.0, 96.0, PixelFormats.Bgr32, null);//定义保存图片大小
        }
        //加载程序
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            _sensor = KinectSensor.GetDefault();

            if (_sensor != null)
            {
                _sensor.Open();

                _reader = _sensor.OpenMultiSourceFrameReader(FrameSourceTypes.Color | FrameSourceTypes.Depth | FrameSourceTypes.Infrared | FrameSourceTypes.Body);
                _reader.MultiSourceFrameArrived += Reader_MultiSourceFrameArrived;
            }
        }
        //结束程序
        private void Window_Closed(object sender, EventArgs e)
        {
            if (_reader != null)
            {
                _reader.Dispose();
            }

            if (_sensor != null)
            {
                _sensor.Close();
            }
        }
        //定义canvas上连线
        protected void DrawingLine(System.Windows.Point startPt, System.Windows.Point endPt)
        {
            LineGeometry myLineGeometry = new LineGeometry();
            myLineGeometry.StartPoint = startPt;
            myLineGeometry.EndPoint = endPt;

            System.Windows.Shapes.Path myPath = new System.Windows.Shapes.Path();
            myPath.Stroke = System.Windows.Media.Brushes.Red;
            myPath.StrokeThickness = 5;
            myPath.Data = myLineGeometry;

            canvas.Children.Add(myPath);

        }
        //定义角度在canvas上面显示
        protected void Addangle(double angle, System.Windows.Point position)
        {
            TextBlock txt = new TextBlock();
            txt.FontSize = 30;
            txt.Foreground = System.Windows.Media.Brushes.White;
            angle = Math.Round(angle, 1);//取小数点后一位
            txt.Text = angle.ToString();
            Canvas.SetTop(txt, position.Y - 50);
            Canvas.SetLeft(txt, position.X - 50);
            canvas.Children.Add(txt);
        }
        //计算关节角度
        public double GetAngle(System.Windows.Point x, System.Windows.Point y, System.Windows.Point z)
        {
            double a = Math.Sqrt((x.Y - y.Y) * (x.Y - y.Y) + (x.X - y.X) * (x.X - y.X));
            double b = Math.Sqrt((z.Y - y.Y) * (z.Y - y.Y) + (z.X - y.X) * (z.X - y.X));
            double c = Math.Sqrt((x.Y - z.Y) * (x.Y - z.Y) + (x.X - z.X) * (x.X - z.X));
            double cosc = (a * a + b * b - c * c) / (2 * a * b);
            double angle = 180 - Math.Acos(cosc) / Math.PI * 180;
            return angle;
        }
        //图像帧叠加
        void ImageCaptured(Bitmap bm)
        {
            if (this.isRecording)
            {
                //这里可能要裁剪
                Bitmap imgRecorded = bm;
                if (this.sizeRevised) // 对图像进行裁剪，  MFile要求录制的视频帧的长和宽必须是4的整数倍。
                {
                    imgRecorded = ESBasic.Helpers.ImageHelper.RoundSizeByNumber(bm, 4);
                    bm.Dispose();
                }                         
                    this.silenceVideoFileMaker.AddVideoFrame(imgRecorded);
            }
        }

        void Reader_MultiSourceFrameArrived(object sender, MultiSourceFrameArrivedEventArgs e)
        {
            var reference = e.FrameReference.AcquireFrame();

            // Color
            using (var frame = reference.ColorFrameReference.AcquireFrame())
            {
                if (frame != null)
                {
                    //将图像转入writeablebitmmap
                    FrameDescription colorFrameDescription = frame.FrameDescription;
                    using (KinectBuffer colorBuffer = frame.LockRawImageBuffer())
                    {
                        this.colorBitmap.Lock();

                        // verify data and write the new color frame data to the display bitmap
                        if ((colorFrameDescription.Width == this.colorBitmap.PixelWidth) && (colorFrameDescription.Height == this.colorBitmap.PixelHeight))
                        {
                            frame.CopyConvertedFrameDataToIntPtr(
                                this.colorBitmap.BackBuffer,
                                (uint)(colorFrameDescription.Width * colorFrameDescription.Height * 4),
                                ColorImageFormat.Bgra);

                            this.colorBitmap.AddDirtyRect(new Int32Rect(0, 0, this.colorBitmap.PixelWidth, this.colorBitmap.PixelHeight));
                        }

                        this.colorBitmap.Unlock();
                    }
                    if (_mode == CameraMode.Color)
                    {
                        camera.Source = frame.ToBitmap();
                    }
                }
            }

            // Depth
            using (var frame = reference.DepthFrameReference.AcquireFrame())
            {
                if (frame != null)
                {
                    if (_mode == CameraMode.Depth)
                    {
                        camera.Source = frame.ToBitmap();
                    }
                }
            }

            // Infrared
            using (var frame = reference.InfraredFrameReference.AcquireFrame())
            {
                if (frame != null)
                {
                    if (_mode == CameraMode.Infrared)
                    {
                        camera.Source = frame.ToBitmap();
                    }
                }
            }

            // Body
            using (var frame = reference.BodyFrameReference.AcquireFrame())
            {
                if (frame != null)
                {
                    canvas.Children.Clear();

                    _bodies = new Body[frame.BodyFrameSource.BodyCount];

                    frame.GetAndRefreshBodyData(_bodies);

                    foreach (var body in _bodies)
                    {
                        if (body.IsTracked)
                        {
                            // COORDINATE MAPPING
                            foreach (Joint joint in body.Joints.Values)
                            {

                                if (joint.TrackingState == TrackingState.Tracked)
                                {
                                    // 3D space point
                                    CameraSpacePoint jointPosition = joint.Position;

                                    // 2D space point
                                    System.Windows.Point point = new System.Windows.Point();


                                    if (_mode == CameraMode.Color)
                                    {
                                        ColorSpacePoint colorPoint = _sensor.CoordinateMapper.MapCameraPointToColorSpace(jointPosition);

                                        point.X = float.IsInfinity(colorPoint.X) ? 0 : colorPoint.X / 2;
                                        point.Y = float.IsInfinity(colorPoint.Y) ? 0 : colorPoint.Y / 2;
                                        //获取肩肘腕，脊椎中间坐标
                                        switch (joint.JointType)
                                        {
                                            case JointType.WristLeft:
                                                WristLeft.X = point.X;
                                                WristLeft.Y = point.Y;
                                                break;
                                            case JointType.ElbowLeft:
                                                ElbowLeft.X = point.X;
                                                ElbowLeft.Y = point.Y;
                                                break;
                                            case JointType.ShoulderLeft:
                                                ShoulderLeft.X = point.X;
                                                ShoulderLeft.Y = point.Y;
                                                break;
                                            case JointType.SpineShoulder:
                                                SpineShoulder.X = point.X;
                                                SpineShoulder.Y = point.Y;
                                                break;
                                            case JointType.ShoulderRight:
                                                ShoulderRight.X = point.X;
                                                ShoulderRight.Y = point.Y;
                                                break;
                                            case JointType.ElbowRight:
                                                ElbowRight.X = point.X;
                                                ElbowRight.Y = point.Y;
                                                break;
                                            case JointType.WristRight:
                                                WristRight.X = point.X;
                                                WristRight.Y = point.Y;
                                                break;
                                        }
                                    }
                                    else if (_mode == CameraMode.Depth || _mode == CameraMode.Infrared) // Change the Image and Canvas dimensions to 512x424
                                    {
                                        DepthSpacePoint depthPoint = _sensor.CoordinateMapper.MapCameraPointToDepthSpace(jointPosition);

                                        point.X = float.IsInfinity(depthPoint.X) ? 0 : depthPoint.X / 2;
                                        point.Y = float.IsInfinity(depthPoint.Y) ? 0 : depthPoint.Y / 2;
                                    }

                                    // Draw 红色圆，直径10
                                    Ellipse ellipse = new Ellipse
                                    {
                                        Fill = System.Windows.Media.Brushes.Red,
                                        Width = 10,
                                        Height = 10
                                    };

                                    Canvas.SetLeft(ellipse, point.X - ellipse.Width / 2);
                                    Canvas.SetTop(ellipse, point.Y - ellipse.Height / 2);
                                    canvas.Children.Add(ellipse);
                                }
                            }


                            if (elbowrightflag)
                            {
                                DrawingLine(ShoulderRight, ElbowRight);//连接肩肘
                                DrawingLine(ElbowRight, WristRight);//连接肘腕

                                //计算肘角度并添加到canvas                                
                                rightelbowangle = GetAngle(ShoulderRight, ElbowRight, WristRight);
                                Addangle(rightelbowangle, ElbowRight);

                            }

                            if (shoulderrightflag)
                            {
                                DrawingLine(ShoulderRight, ElbowRight);//连接肩肘
                                DrawingLine(SpineShoulder, ShoulderRight);//连接肩和肩中
                                //计算肩角度并添加到canvas
                                rightshoulderangle = GetAngle(SpineShoulder, ShoulderRight, ElbowRight);
                                Addangle(rightshoulderangle, ShoulderRight);
                            }

                        }
                    }
                }
            }
        }

        private void shoulderangle_Checked(object sender, RoutedEventArgs e)
        {
            shoulderrightflag = true;
        }

        private void shoulderangle_Unchecked(object sender, RoutedEventArgs e)
        {
            shoulderrightflag = false;
        }

        private void elbowangle_Checked(object sender, RoutedEventArgs e)
        {
            elbowrightflag = true;
        }

        private void elbowangle_Unchecked(object sender, RoutedEventArgs e)
        {
            elbowrightflag = false;
        }

        private void screenshot_Click(object sender, RoutedEventArgs e)
        {
            if (this.colorBitmap != null)
            {
                // create a png bitmap encoder which knows how to save a .png file
                BitmapEncoder encoder = new PngBitmapEncoder();

                // create frame from the writable bitmap and add to encoder
                encoder.Frames.Add(BitmapFrame.Create(this.colorBitmap));

                string time = System.DateTime.Now.ToString("hh'-'mm'-'ss", CultureInfo.CurrentUICulture.DateTimeFormat);

                string myPhotos = "E:\\testDir\\";

                string path = System.IO.Path.Combine(myPhotos, "KinectScreenshot-Color-" + time + ".png");

                // write the new file to disk
                try
                {
                    // FileStream is IDisposable
                    using (FileStream fs = new FileStream(path, FileMode.Create))
                    {
                        encoder.Save(fs);
                    }

                    //this.StatusText = string.Format(Properties.Resources.SavedScreenshotStatusTextFormat, path);
                }
                catch (IOException)
                {
                    //this.StatusText = string.Format(Properties.Resources.FailedScreenshotStatusTextFormat, path);
                }

                //写入关节角度数据
                //如果文件不存在，则创建；存在则覆盖
                //该方法写入字符数组换行显示
                string[] lines = { time, "elbowangle:", rightelbowangle.ToString(), "shoulderangle:", rightshoulderangle.ToString(), "\n" };
                //System.IO.File.WriteAllLines(@"E:\testDir\test.txt", lines, Encoding.UTF8);

                //StreamWriter一个参数默认覆盖
                //StreamWriter第二个参数为false覆盖现有文件，为true则把文本追加到文件末尾
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"E:\testDir\test.txt", true))
                {
                    foreach (string line in lines)
                    {
                        //file.Write(line);//直接追加文件末尾，不换行
                        file.WriteLine(line);// 直接追加文件末尾，换行  

                    }
                }
            }
        }
        private void start_Click(object sender, RoutedEventArgs e)
        {
            this.isRecording = true;
            System.Drawing.Size videoSize = Screen.PrimaryScreen.Bounds.Size;
            System.Drawing.Rectangle myRect = new System.Drawing.Rectangle();
            RECT rct = new RECT();
            GetWindowRect(new HandleRef(this, new WindowInteropHelper(this).Handle), out rct);
            myRect.X = rct.Left;
            myRect.Y = rct.Top;
            myRect.Width = rct.Right - rct.Left + 1;
            myRect.Height = rct.Bottom - rct.Top + 1;
            this.desktopCapturer = CapturerFactory.CreateDesktopCapturer(frameRate, false,myRect);
            this.desktopCapturer.ImageCaptured += ImageCaptured;
            videoSize = this.desktopCapturer.VideoSize;
            this.desktopCapturer.Start();
            this.start.IsEnabled = false;

            this.sizeRevised = (videoSize.Width % 4 != 0) || (videoSize.Height % 4 != 0);
            if (this.sizeRevised)
            {
                videoSize = new System.Drawing.Size(videoSize.Width / 4 * 4, videoSize.Height / 4 * 4);
            }
            string time = System.DateTime.Now.ToString("hh'-'mm'-'ss", CultureInfo.CurrentUICulture.DateTimeFormat);
            this.silenceVideoFileMaker = new SilenceVideoFileMaker();
            this.silenceVideoFileMaker.Initialize("E:\\testDir\\"+time+".mp4", VideoCodecType.H264, videoSize.Width, videoSize.Height, frameRate, VideoQuality.High);
        }

        private void end_Click(object sender, RoutedEventArgs e)
        {
            this.start.IsEnabled = true;
            this.isRecording = false;
            this.desktopCapturer.Stop();
            this.silenceVideoFileMaker.Close(true);
        }
    }



    enum CameraMode
    {
        Color,
        Depth,
        Infrared
    }
}
