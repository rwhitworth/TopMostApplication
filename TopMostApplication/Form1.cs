using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using IniFile;
using System.Drawing.Imaging;

namespace TopMostApplication
{
    public partial class Form1 : Form
    {

        #region Win32 interop 
        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        [DllImport("user32.dll")]
        static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;        // x position of upper-left corner
            public int Top;         // y position of upper-left corner
            public int Right;       // x position of lower-right corner
            public int Bottom;      // y position of lower-right corner
        }

        [DllImport("user32.dll")]
        static extern bool SetActiveWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        private string GetActiveWindowTitle()
        {
            const int nChars = 256;
            StringBuilder Buff = new StringBuilder(nChars);
            IntPtr handle = GetForegroundWindow();

            if (GetWindowText(handle, Buff, nChars) > 0)
            {
                return Buff.ToString();
            }
            return null;
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern int SendMessage(IntPtr hwnd, int wMsg, IntPtr wParam, IntPtr lParam);
        #endregion

        private RECT previous_location = new RECT();
        private bool should_hide_button = false;
        private IntPtr IE_IntPtr = IntPtr.Zero;
        private bool screenshot_in_action = false;
        private System.Drawing.Imaging.ImageFormat imageformat = System.Drawing.Imaging.ImageFormat.Jpeg;
        private string imageext = "jpg";
        private Keys hotkey = Keys.F10;
        private string ProgramTitle = "Internet Explorer".ToUpper();
        private long JPEGQuality = 100;
        private string FileNameTemplate = "yyyyMMdd-HHmmss-ff";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            TransparencyKey = Color.FromArgb(255, 255, 255, 255);
            BackColor = Color.FromArgb(255, 255, 255, 255);
            button1.Hide();
            timer_trackvisible.Interval = 100;
            timer_trackmovement.Interval = 1;

            #region read ini settings

            var MyIni = new IniFile.IniFile("TopMostApplication.exe.config");

            if (MyIni.KeyExists("ButtonText", "ini"))
            {
                string buttontext = MyIni.Read("ButtonText", "ini");
                button1.Text = buttontext;
            }

            if (MyIni.KeyExists("ImageFormat", "ini"))
            {
                string imageformatstr = MyIni.Read("ImageFormat", "ini").ToUpper();
                switch (imageformatstr)
                {
                    case "JPEG": imageformat = System.Drawing.Imaging.ImageFormat.Jpeg; imageext = "jpg";  break;
                    case "JPG": imageformat = System.Drawing.Imaging.ImageFormat.Jpeg; imageext = "jpg"; break;
                    case "PNG": imageformat = System.Drawing.Imaging.ImageFormat.Png; imageext = "png"; break;
                    case "BMP": imageformat = System.Drawing.Imaging.ImageFormat.Bmp; imageext = "bmp"; break;
                    case "TIFF": imageformat = System.Drawing.Imaging.ImageFormat.Tiff; imageext = "tif"; break;
                    case "TIF": imageformat = System.Drawing.Imaging.ImageFormat.Tiff; imageext = "tif"; break;
                    default:
                        break;
                }
            }

            if (MyIni.KeyExists("HotKey", "ini"))
            {
                string hotkeystr = MyIni.Read("HotKey", "ini").ToUpper();
                switch (hotkeystr)
                {
                    case "F1": hotkey = Keys.F1; break;
                    case "F2": hotkey = Keys.F2; break;
                    case "F3": hotkey = Keys.F3; break;
                    case "F4": hotkey = Keys.F4; break;
                    case "F5": hotkey = Keys.F5; break;
                    case "F6": hotkey = Keys.F6; break;
                    case "F7": hotkey = Keys.F7; break;
                    case "F8": hotkey = Keys.F8; break;
                    case "F9": hotkey = Keys.F9; break;
                    case "F10": hotkey = Keys.F10; break;
                    case "F11": hotkey = Keys.F11; break;
                    case "F12": hotkey = Keys.F12; break;
                    default:
                        break;
                }
            }

            if (MyIni.KeyExists("ProgramTitle", "ini"))
            {
                ProgramTitle = MyIni.Read("ProgramTitle", "ini").ToUpper();
            }

            if (MyIni.KeyExists("JPEGQuality", "ini"))
            {
                bool result = long.TryParse(MyIni.Read("JPEGQuality", "ini"), out JPEGQuality);
                if (!result) { JPEGQuality = 100; }
            }

            if (MyIni.KeyExists("FileNameTemplate", "ini"))
            {
                FileNameTemplate = MyIni.Read("FileNameTemplate", "ini");
                if (FileNameTemplate == null) { FileNameTemplate = "yyyyMMdd-HHmmss-ff"; }
            }

            #endregion

            hotkey.Hotkey h = new hotkey.Hotkey(hotkey, true, false, false, false);
            h.Pressed += H_Pressed;
            h.Register(this);

            // and finally, enable the timers
            // do this last so all configuration is complete before they can possibly run
            timer_trackvisible.Enabled = true;
            timer_trackmovement.Enabled = true;
        }

        private void H_Pressed(object sender, HandledEventArgs e)
        {
            take_screenshot();
        }

        private string ConvertDateString(string DateStr)
        {
            DateTime dt = DateTime.Now;
            return dt.ToString(DateStr);
        }

        private void take_screenshot()
        {
            if (IE_IntPtr != IntPtr.Zero)
            {
                screenshot_in_action = true;
                button1.Hide();
                SetForegroundWindow(IE_IntPtr);
                RECT r = new RECT();
                bool result = GetWindowRect(IE_IntPtr, out r);

                using (Bitmap bitmap = new Bitmap(r.Right - r.Left, r.Bottom - r.Top))
                {
                    using (Graphics g = Graphics.FromImage(bitmap))
                    {
                        g.CopyFromScreen(new Point(r.Left, r.Top), Point.Empty, new Size(bitmap.Size.Width, bitmap.Size.Height));
                    }
                    //bitmap.Save("test." + imageext, imageformat);
                    SaveJpegWithCompression(bitmap, ConvertDateString(FileNameTemplate) + "." + imageext, JPEGQuality);
                }
            }

            screenshot_in_action = false;
            button1.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            take_screenshot();
        }

        private void timer_trackmovement_Tick(object sender, EventArgs e)
        {
            RECT current_location = new RECT();
            bool result = GetWindowRect(GetForegroundWindow(), out current_location);
            if (result)
            {
                if ((current_location.Left != previous_location.Left) || (current_location.Right != previous_location.Right)
                    || (current_location.Top != previous_location.Top) || (current_location.Bottom != previous_location.Bottom))
                {
                    previous_location = current_location;
                    SetDesktopLocation(current_location.Left, current_location.Top);
                }
            }
        }

        private void timer_trackvisible_Tick(object sender, EventArgs e)
        {
            System.Threading.Thread.Sleep(1);
            string activewindowtitle = GetActiveWindowTitle();
            if (activewindowtitle == null)
            {
                activewindowtitle = "";
            }
            if (activewindowtitle.ToUpper().Contains(ProgramTitle))
            {
                IE_IntPtr = GetForegroundWindow();
                if (!screenshot_in_action) { button1.Show(); }
                should_hide_button = false;
                RECT r = new RECT();
                bool result = GetWindowRect(GetForegroundWindow(), out r);
                if (result)
                {
                    SetDesktopLocation(r.Left, r.Top);
                }
            }
            else if (should_hide_button)
            {
                button1.Hide();
            }
            else
            {
                should_hide_button = true;
            }

        }

        #region JPEG code
        /// <summary>
        /// Retrieves the Encoder Information for a given MimeType
        /// </summary>
        /// <param name="mimeType">String: Mimetype</param>
        /// <returns>ImageCodecInfo: Mime info or null if not found</returns>
        private static ImageCodecInfo GetEncoderInfo(String mimeType)
        {
            var encoders = ImageCodecInfo.GetImageEncoders();
            return encoders.FirstOrDefault(t => t.MimeType == mimeType);
        }

        /// <summary>
        /// Save an Image as a JPeg with a given compression
        ///  Note: Filename suffix will not affect mime type which will be Jpeg.
        /// </summary>
        /// <param name="image">Image: Image to save</param>
        /// <param name="fileName">String: File name to save the image as. Note: suffix will not affect mime type which will be Jpeg.</param>
        /// <param name="compression">Long: Value between 0 and 100.</param>
        private static void SaveJpegWithCompressionSetting(Image image, string fileName, long compression)
        {
            var eps = new EncoderParameters(1);
            eps.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, compression);
            var ici = GetEncoderInfo("image/jpeg");
            image.Save(fileName, ici, eps);
        }

        /// <summary>
        /// Save an Image as a JPeg with a given compression
        /// Note: Filename suffix will not affect mime type which will be Jpeg.
        /// </summary>
        /// <param name="image">Image: This image</param>
        /// <param name="fileName">String: File name to save the image as. Note: suffix will not affect mime type which will be Jpeg.</param>
        /// <param name="compression">Long: Value between 0 and 100.</param>
        public static void SaveJpegWithCompression(Image image, string fileName, long compression)
        {
            SaveJpegWithCompressionSetting(image, fileName, compression);
        }
        public static void SaveJpegWithCompression(Bitmap bitmap, string fileName, long compression)
        {
            var eps = new EncoderParameters(1);
            eps.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, compression);
            var ici = GetEncoderInfo("image/jpeg");
            bitmap.Save(fileName, ici, eps);
        }
        #endregion
    }
}
