using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SurfaceAutomation
{
    public class clsResolution
    {
        [DllImport("gdi32.dll", CharSet = CharSet.Auto, SetLastError = true, ExactSpelling = true)]
        public static extern int GetDeviceCaps(IntPtr hDC, int nIndex);

        public enum DeviceCap
        {
            VERTRES = 10,
            DESKTOPVERTRES = 117
        }

        public enum EXECUTION_STATE : uint
        {
            ES_AWAYMODE_REQUIRED = 0x00000040,
            ES_CONTINUOUS = 0x80000000,
            ES_DISPLAY_REQUIRED = 0x00000002,
            ES_SYSTEM_REQUIRED = 0x00000001
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern EXECUTION_STATE setThreadExecutionState(EXECUTION_STATE esFlags);

        public void avoidlock()
        {
            setThreadExecutionState(EXECUTION_STATE.ES_DISPLAY_REQUIRED | EXECUTION_STATE.ES_SYSTEM_REQUIRED);
            Thread.Sleep(30000);
            avoidlock();
        }
        public static double GetWindowsScreenScalingFactor(bool percentage = true)
        {
            //Create Graphics object from the current windows handle
            Graphics GraphicsObject = Graphics.FromHwnd(IntPtr.Zero);
            //Get Handle to the device context associated with this Graphics object
            IntPtr DeviceContextHandle = GraphicsObject.GetHdc();
            //Call GetDeviceCaps with the Handle to retrieve the Screen Height
            int LogicalScreenHeight = GetDeviceCaps(DeviceContextHandle, (int)DeviceCap.VERTRES);
            int PhysicalScreenHeight = GetDeviceCaps(DeviceContextHandle, (int)DeviceCap.DESKTOPVERTRES);
            //Divide the Screen Heights to get the scaling factor and round it to two decimals
            double ScreenScalingFactor = Math.Round(PhysicalScreenHeight / (double)LogicalScreenHeight, 2);
            //If requested as percentage - convert it
            if (percentage)
            {
                ScreenScalingFactor *= 100.0;
            }
            //Release the Handle and Dispose of the GraphicsObject object
            GraphicsObject.ReleaseHdc(DeviceContextHandle);
            GraphicsObject.Dispose();
            //Return the Scaling Factor
            return ScreenScalingFactor;
        }

        public static Size GetDisplayResolution()
        {
            var sf = GetWindowsScreenScalingFactor(false);
            var screenWidth = Screen.PrimaryScreen.Bounds.Width * sf;
            var screenHeight = Screen.PrimaryScreen.Bounds.Height * sf;
            return new Size((int)screenWidth, (int)screenHeight);
        }

        public struct DEVMODE1
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
            public string dmDeviceName;
            public short dmSpecVersion;
            public short dmDriverVersion;
            public short dmSize;
            public short dmDriverExtra;
            public int dmFields;

            public short dmOrientation;
            public short dmPaperSize;
            public short dmPaperLength;
            public short dmPaperWidth;

            public short dmScale;
            public short dmCopies;
            public short dmDefaultSource;
            public short dmPrintQuality;
            public short dmColor;
            public short dmDuplex;
            public short dmYResolution;
            public short dmTTOption;
            public short dmCollate;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
            public string dmFormName;
            public short dmLogPixels;
            public short dmBitsPerPel;
            public int dmPelsWidth;
            public int dmPelsHeight;

            public int dmDisplayFlags;
            public int dmDisplayFrequency;

            public int dmICMMethod;
            public int dmICMIntent;
            public int dmMediaType;
            public int dmDitherType;
            public int dmReserved1;
            public int dmReserved2;

            public int dmPanningWidth;
            public int dmPanningHeight;

        }

        public class User_32
        {
            [DllImport("user32.dll")]
            public static extern int EnumDisplaySettings(string deviceName, int modeNum, ref DEVMODE1 devmode);
            [DllImport("User_32.dll")]
            public static extern int ChangeDisplaySettings(ref DEVMODE1 devmode, int flags);

            public const int ENUM_CURRENT_SETTINGS = -1;
            public const int CDS_UPDATEREGISTRY = 0x01;
            public const int CDS_TEST = 0x02;
            public const int DISP_CHANGE_SUCCESSFUL = 0;
            public const int DISP_CHANGE_RESTART = 1;
            public const int DISP_CHANGE_FAILED = -1;
        }

        [DllImport("user32.dll", SetLastError = true)]
        static extern void SwitchToThisWindow(IntPtr hWnd, bool FaltTab);
        public void FindNoOfOpenWindow()
        {
            int i = 0;
            int j = 0;
            int chck = 0;
            string[] arrayWindowTitle = new string[20];
            Process[] process = Process.GetProcesses();
            foreach(Process p in process)
            {
                if(!String.IsNullOrEmpty(p.MainWindowTitle))
                {
                    arrayWindowTitle[i] = p.MainWindowTitle.ToString();
                    if(p.MainWindowTitle.ToString() == "OpicsPlus - \\\\Remote" || p.MainWindowTitle.ToString() == "LOGR - \\\\Remote")
                    {
                        SwitchToThisWindow(p.MainWindowHandle, true);
                        chck = 1;
                        break;
                    }
                    i++;
                }
                j++;
            }
        }
        public class CResolution
        {
            public CResolution(int a, int b)
            {
                Screen scrn = Screen.PrimaryScreen;
                int iWidth = a;
                int iHeight = b;

                DEVMODE1 dm = new DEVMODE1();
                dm.dmDeviceName = new String(new char[32]);
                dm.dmFormName = new String(new char[32]);
                dm.dmSize = (short)Marshal.SizeOf(dm);

                if(0 != User_32.EnumDisplaySettings(null, User_32.ENUM_CURRENT_SETTINGS, ref dm))
                {
                    dm.dmPelsHeight = iHeight;
                    dm.dmPelsWidth = iWidth;

                    int iRet = User_32.ChangeDisplaySettings(ref dm, User_32.CDS_TEST);

                    if(iRet == User_32.DISP_CHANGE_FAILED)
                    {
                        Console.WriteLine("Failed to change the solution");
                    }
                    else
                    {
                        iRet = User_32.ChangeDisplaySettings(ref dm, User_32.CDS_UPDATEREGISTRY);

                        switch (iRet)
                        {
                            case User_32.DISP_CHANGE_SUCCESSFUL:
                                break;
                            case User_32.DISP_CHANGE_RESTART:
                                Console.WriteLine("You need to Reboot for the change to happen.\n If you feel any problem after rebooting your machine\n Then try to change resolution in safe mode.");
                                break;
                            default:
                                Console.WriteLine("Failed to change the solution");
                                break;
                        }
                    }
                }
            }
        }
    }
}
