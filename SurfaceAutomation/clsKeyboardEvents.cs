using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Threading;

namespace SurfaceAutomation
{
    public class clsKeyboardEvents
    {
        [DllImport("user32.dll")]
        public static extern int FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        public static extern IntPtr SetForegroundWidow(int hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        static extern void keybd_event(int bvK, int bScan, int dwFlags, int dwExtraInfo);

        [DllImport("user32.dll")]
        public static extern IntPtr GetDC(IntPtr hwnd);

        [DllImport("user32.dll")]
        public static extern IntPtr ReleaseDC(IntPtr hwnd, IntPtr dc);

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);


        private const int MOUSEEVENTIF_LEFTDOWN = 0x02;
        private const int MOUSEEVENTIF_LEFTUP = 0x02;
        private const int MOUSEEVENTIF_RIGHTDOWN = 0x02;
        private const int MOUSEEVENTIF_RIGHTUP = 0x02;

        public const int KEYEVENTIF_EXTENDEDKEY = 0x0001;
        public const int KEYEVENTIF_KEYUP = 0x0002;

        public const int MOUSEEVENTIF_MIDDLEDOWN = 0x20;
        public const int MOUSEEVENTIF_MIDDLEUP = 0x40;

        public const int VK_LCONTROL = 0xA2;
        public const int VK_SHIFT = 16;
        public const int VK_TAB = 9;
        public const int VK_RETURN = 13;
        public const int VK_MENU = 17;
        public const int VK_PAUSE = 18;
        public const int VK_CAPITAL = 19;
        public const int VK_NUMLOCK = 20;
        public const int VK_SCROLL = 21;
        public const int VK_EXECUTE = 22;
        public const int VK_INSERT = 23;
        public const int VK_DELETE = 24;
        public const int VK_HELP = 25;

        public bool shiftkey;
        public bool ctrlkey;
        public bool altkey;
        public bool shift;
        public bool ctrl;
        public bool alt;

        public int keyval;

        public void keyInText(string inputString)
        {
            char[] array = inputString.ToCharArray();
            foreach(char c in array)
            {
                shiftkey = false;
                ctrlkey = false;
                switch (c)
                {
                    case '!': shiftkey = true; ctrlkey = false; keyval = 33; break;
                    case '"': shiftkey = true; ctrlkey = false; keyval = 34; break;
                    case '#': shiftkey = true; ctrlkey = false; keyval = 35; break;
                    case '$': shiftkey = true; ctrlkey = false; keyval = 36; break;
                    case '%': shiftkey = true; ctrlkey = false; keyval = 37; break;
                    case '&': shiftkey = true; ctrlkey = false; keyval = 38; break;
                    case '\'': shiftkey = false; ctrlkey = false; keyval=39; break;
                    case '(': shiftkey = true; ctrlkey = false; keyval = 40; break;
                    case ')': shiftkey = true; ctrlkey = false; keyval = 41; break;
                    case '*': shiftkey = true; ctrlkey = false; keyval = 42; break;
                    case '+': shiftkey = true; ctrlkey = false; keyval = 43; break;
                    case ',': shiftkey = false; ctrlkey = false; keyval = 44; break;
                    case '-': shiftkey = false; ctrlkey = false; keyval = 45; break;
                    case '.': shiftkey = false; ctrlkey = false; keyval = 46; break;
                    case '/': shiftkey = false; ctrlkey = false; keyval = 47; break;
                    case '0': shiftkey = false; ctrlkey = false; keyval = 48; break;
                    case '1': shiftkey = false; ctrlkey = false; keyval = 49; break;
                    case '2': shiftkey = false; ctrlkey = false; keyval = 50; break;
                    case '3': shiftkey = false; ctrlkey = false; keyval = 51; break;
                    case '4': shiftkey = false; ctrlkey = false; keyval = 52; break;
                    case '5': shiftkey = false; ctrlkey = false; keyval = 53; break;
                    case '6': shiftkey = false; ctrlkey = false; keyval = 54; break;
                    case '7': shiftkey = false; ctrlkey = false; keyval = 55; break;
                    case '8': shiftkey = false; ctrlkey = false; keyval = 56; break;
                    case '9': shiftkey = false; ctrlkey = false; keyval = 57; break;
                    case ':': shiftkey = true; ctrlkey = false; keyval = 58; break;
                    case ';': shiftkey = false; ctrlkey = false; keyval = 59; break; 
                    case '<': shiftkey = true; ctrlkey = false; keyval = 60; break;
                    case '=': shiftkey = true; ctrlkey = false; keyval = 61; break;
                    case '>': shiftkey = true; ctrlkey = false; keyval = 62; break;
                    case '?': shiftkey = true; ctrlkey = false; keyval = 63; break;
                    case '@': shiftkey = true; ctrlkey = false; keyval = 64; break;
                    case 'A': shiftkey = true; ctrlkey = false; keyval = 65; break;
                    case 'B': shiftkey = true; ctrlkey = false; keyval = 66; break;
                    case 'C': shiftkey = true; ctrlkey = false; keyval = 67; break;
                    case 'D': shiftkey = true; ctrlkey = false; keyval = 68; break;
                    case 'E': shiftkey = true; ctrlkey = false; keyval = 69; break;
                    case 'F': shiftkey = true; ctrlkey = false; keyval = 70; break;
                    case 'G': shiftkey = true; ctrlkey = false; keyval = 71; break;
                    case 'H': shiftkey = true; ctrlkey = false; keyval = 72; break;
                    case 'I': shiftkey = true; ctrlkey = false; keyval = 73; break;
                    case 'J': shiftkey = true; ctrlkey = false; keyval = 74; break;
                    case 'K': shiftkey = true; ctrlkey = false; keyval = 75; break;
                    case 'L': shiftkey = true; ctrlkey = false; keyval = 76; break;
                    case 'M': shiftkey = true; ctrlkey = false; keyval = 77; break;
                    case 'N': shiftkey = true; ctrlkey = false; keyval = 78; break;
                    case 'O': shiftkey = true; ctrlkey = false; keyval = 79; break;
                    case 'P': shiftkey = true; ctrlkey = false; keyval = 80; break;
                    case 'Q': shiftkey = true; ctrlkey = false; keyval = 81; break;
                    case 'R': shiftkey = true; ctrlkey = false; keyval = 82; break;
                    case 'S': shiftkey = true; ctrlkey = false; keyval = 83; break;
                    case 'T': shiftkey = true; ctrlkey = false; keyval = 84; break;
                    case 'U': shiftkey = true; ctrlkey = false; keyval = 85; break;
                    case 'V': shiftkey = true; ctrlkey = false; keyval = 86; break;
                    case 'W': shiftkey = true; ctrlkey = false; keyval = 87; break;
                    case 'X': shiftkey = true; ctrlkey = false; keyval = 88; break;
                    case 'Y': shiftkey = true; ctrlkey = false; keyval = 89; break;
                    case 'Z': shiftkey = true; ctrlkey = false; keyval = 90; break;
                    case '[': shiftkey = false; ctrlkey = false; keyval = 91; break; 
                    case '\\': shiftkey = false; ctrlkey = false; keyval = 92; break;
                    case ']': shiftkey = false; ctrlkey = false; keyval = 93; break;
                    case '^': shiftkey = true; ctrlkey = false; keyval = 94; break;
                    case '_': shiftkey = true; ctrlkey = false; keyval = 95; break;
                    case '`': shiftkey = false; ctrlkey = false; keyval = 96; break;
                    case 'a': shiftkey = false; ctrlkey = false; keyval = 97; break;
                    case 'b': shiftkey = false; ctrlkey = false; keyval = 98; break;
                    case 'c': shiftkey = false; ctrlkey = false; keyval = 99; break;
                    case 'd': shiftkey = false; ctrlkey = false; keyval = 100; break;
                    case 'e': shiftkey = false; ctrlkey = false; keyval = 101; break;
                    case 'f': shiftkey = false; ctrlkey = false; keyval = 102; break;
                    case 'g': shiftkey = false; ctrlkey = false; keyval = 103; break;
                    case 'h': shiftkey = false; ctrlkey = false; keyval = 104; break;
                    case 'i': shiftkey = false; ctrlkey = false; keyval = 105; break;
                    case 'j': shiftkey = false; ctrlkey = false; keyval = 106; break;
                    case 'k': shiftkey = false; ctrlkey = false; keyval = 107; break;
                    case 'l': shiftkey = false; ctrlkey = false; keyval = 108; break;
                    case 'm': shiftkey = false; ctrlkey = false; keyval = 109; break;
                    case 'n': shiftkey = false; ctrlkey = false; keyval = 110; break;
                    case 'o': shiftkey = false; ctrlkey = false; keyval = 111; break;
                    case 'p': shiftkey = false; ctrlkey = false; keyval = 112; break;
                    case 'q': shiftkey = false; ctrlkey = false; keyval = 113; break;
                    case 'r': shiftkey = false; ctrlkey = false; keyval = 114; break;
                    case 's': shiftkey = false; ctrlkey = false; keyval = 115; break;
                    case 't': shiftkey = false; ctrlkey = false; keyval = 116; break;
                    case 'u': shiftkey = false; ctrlkey = false; keyval = 117; break;
                    case 'v': shiftkey = false; ctrlkey = false; keyval = 118; break;
                    case 'w': shiftkey = false; ctrlkey = false; keyval = 119; break;
                    case 'x': shiftkey = false; ctrlkey = false; keyval = 120; break;
                    case 'y': shiftkey = false; ctrlkey = false; keyval = 121; break;
                    case 'z': shiftkey = false; ctrlkey = false; keyval = 122; break;
                    case '{': shiftkey = true; ctrlkey = false; keyval = 123; break;
                    case '|': shiftkey = true; ctrlkey = false; keyval = 124; break;
                    case '}': shiftkey = true; ctrlkey = false; keyval = 125; break;
                    case '~': shiftkey = true; ctrlkey = false; keyval = 126; break;
                    default:break;
                }
                if (shiftkey)
                {
                    keybd_event(VK_SHIFT, 170, KEYEVENTIF_EXTENDEDKEY, 0);
                    keybd_event(keyval, 170, KEYEVENTIF_EXTENDEDKEY, 0);
                    keybd_event(keyval, 170, KEYEVENTIF_KEYUP, 0);
                    keybd_event(VK_SHIFT, 170, KEYEVENTIF_KEYUP, 0);
                    Thread.Sleep(10);
                }
                else
                {
                    keybd_event(keyval, 0, KEYEVENTIF_EXTENDEDKEY, 0);
                    keybd_event(keyval, 0, KEYEVENTIF_KEYUP, 0);
                    Thread.Sleep(10);
                }
            }
        }
        
        public void StartProcess(int x_axis, int y_axis, int ht, int wd,int wait, string action, string query)
        {
            switch(action)
            {
                case "passvalue":
                    keyInText(query);
                    break;
                default:
                    ProcessExecution(action, x_axis, y_axis, wd, 0, wait);
                    break;
            }
        }

        public void ProcessExecution(string action, int x_axis, int y_axis, int wd, int ht, int wait)
        {
            switch(action)
            {
                case "LCLICK":
                    mouse_event(MOUSEEVENTIF_LEFTDOWN, x_axis, y_axis, 0, 0);
                    mouse_event(MOUSEEVENTIF_LEFTUP, x_axis, y_axis, 0, 0);
                    Thread.Sleep(wait);
                    break;
                case "RCLICK":
                    mouse_event(MOUSEEVENTIF_RIGHTDOWN, x_axis, y_axis, 0, 0);
                    mouse_event(MOUSEEVENTIF_RIGHTUP, x_axis, y_axis, 0, 0);
                    Thread.Sleep(wait);
                    break;
                case "TAB":
                    keybd_event(0x09, 0x0f, KEYEVENTIF_EXTENDEDKEY, 0);
                    keybd_event(0x09, 0, KEYEVENTIF_KEYUP, 0);
                    Thread.Sleep(wait);
                    break;
                case "F15":
                    keybd_event(126, 0x0f, KEYEVENTIF_EXTENDEDKEY, 0);
                    keybd_event(126, 0, KEYEVENTIF_KEYUP, 0);
                    Thread.Sleep(wait);
                    break;
                case "PRINTSCREEN":
                    keybd_event(44, 0x0f, 0, 0);
                    keybd_event(44, 0, KEYEVENTIF_KEYUP, 0);
                    Thread.Sleep(wait);
                    break;
                case "NUMLOCK":
                    keybd_event(0xC5, 45, KEYEVENTIF_EXTENDEDKEY, 0);
                    keybd_event(0xC5, 45, KEYEVENTIF_KEYUP, 0);
                    Thread.Sleep(wait);
                    break;
                case "ALT":
                    keybd_event(0xB8, 0x38, KEYEVENTIF_EXTENDEDKEY, 0);
                    keybd_event(0xB8, 0x38, KEYEVENTIF_KEYUP, 0);
                    Thread.Sleep(wait);
                    break;
                case "ALTDOWN+":
                    keybd_event(0xB8, 0x38, KEYEVENTIF_EXTENDEDKEY, 0);
                    Thread.Sleep(wait);
                    break;
                case "ALTUP+":
                    keybd_event(0xB8, 0x38, KEYEVENTIF_KEYUP, 0);
                    Thread.Sleep(wait);
                    break;
                case "CTRL":
                    keybd_event(0x9D, 0x1D, KEYEVENTIF_EXTENDEDKEY, 0);
                    keybd_event(0x9D, 0x1D, KEYEVENTIF_KEYUP, 0);
                    Thread.Sleep(wait);
                    break;
                case "CTRLDOWN+":
                    keybd_event(0x9D, 0x1D, KEYEVENTIF_EXTENDEDKEY, 0);
                    Thread.Sleep(wait);
                    break;
                case "CTRLUP+":
                    keybd_event(0x9D, 0x1D, KEYEVENTIF_KEYUP, 0);
                    Thread.Sleep(wait);
                    break;
                case "SHIFT":
                    keybd_event(0xAA, 0x2A, KEYEVENTIF_EXTENDEDKEY, 0);
                    keybd_event(0xAA, 0x2A, KEYEVENTIF_KEYUP, 0);
                    Thread.Sleep(wait);
                    break;
                case "SHIFTDOWN+":
                    keybd_event(0xAA, 0x2A, KEYEVENTIF_EXTENDEDKEY, 0);
                    Thread.Sleep(wait);
                    break;
                case "SHIFTUP+":
                    keybd_event(0xAA, 0x2A, KEYEVENTIF_KEYUP, 0);
                    Thread.Sleep(wait);
                    break;
                case "ENTER":
                    keybd_event(0x9C, 0x1C, KEYEVENTIF_EXTENDEDKEY, 0);
                    keybd_event(0x9C, 0x1C, KEYEVENTIF_KEYUP, 0);
                    Thread.Sleep(wait);
                    break;
                case "SPACE":
                    keybd_event(32, 0, KEYEVENTIF_EXTENDEDKEY, 0);
                    keybd_event(32, 0, KEYEVENTIF_KEYUP, 0);
                    Thread.Sleep(wait);
                    break;
                case "DELETE":
                    keybd_event(46, 0, KEYEVENTIF_EXTENDEDKEY, 0);
                    keybd_event(46, 0, KEYEVENTIF_KEYUP, 0);
                    Thread.Sleep(wait);
                    break;
                case "BACKSPACE":
                    keybd_event(8, 0, KEYEVENTIF_EXTENDEDKEY, 0);
                    keybd_event(8, 0, KEYEVENTIF_KEYUP, 0);
                    Thread.Sleep(wait);
                    break;
                case "PAGEUP":
                    keybd_event(33, 0, KEYEVENTIF_EXTENDEDKEY, 0);
                    keybd_event(33, 0, KEYEVENTIF_KEYUP, 0);
                    Thread.Sleep(wait);
                    break;
                case "PAGEDOWN":
                    keybd_event(34, 0, KEYEVENTIF_EXTENDEDKEY, 0);
                    keybd_event(34, 0, KEYEVENTIF_KEYUP, 0);
                    Thread.Sleep(wait);
                    break;
                case "HOME":
                    keybd_event(36, 0, KEYEVENTIF_EXTENDEDKEY, 0);
                    keybd_event(36, 0, KEYEVENTIF_KEYUP, 0);
                    Thread.Sleep(wait);
                    break;
                case "END":
                    keybd_event(35, 0, KEYEVENTIF_EXTENDEDKEY, 0);
                    keybd_event(35, 0, KEYEVENTIF_KEYUP, 0);
                    Thread.Sleep(wait);
                    break;
                case "CUT":
                    keybd_event(0x9D, 0x1D, KEYEVENTIF_EXTENDEDKEY, 0);
                    keybd_event(0xAD, 0x2D, KEYEVENTIF_EXTENDEDKEY, 0);
                    keybd_event(0xAD, 0x2D, KEYEVENTIF_KEYUP, 0);
                    keybd_event(0x9D, 0x1D, KEYEVENTIF_KEYUP, 0);
                    Thread.Sleep(wait);
                    break;
                case "COPY":
                    keybd_event(0x9D, 0x1D, KEYEVENTIF_EXTENDEDKEY, 0);
                    keybd_event(0xAE, 0x2D, KEYEVENTIF_EXTENDEDKEY, 0);
                    keybd_event(0xAE, 0x2D, KEYEVENTIF_KEYUP, 0);
                    keybd_event(0x9D, 0x1D, KEYEVENTIF_KEYUP, 0);
                    Thread.Sleep(wait);
                    break;
                case "PASTE":
                    keybd_event(0x9D, 0x1D, KEYEVENTIF_EXTENDEDKEY, 0);
                    keybd_event(0xAF, 0x2D, KEYEVENTIF_EXTENDEDKEY, 0);
                    keybd_event(0xAF, 0x2D, KEYEVENTIF_KEYUP, 0);
                    keybd_event(0x9D, 0x1D, KEYEVENTIF_KEYUP, 0);
                    Thread.Sleep(wait);
                    break;
                case "SELECTALL":
                    keybd_event(0x9D, 0x1D, KEYEVENTIF_EXTENDEDKEY, 0);
                    keybd_event(0x9e, 0x1e, KEYEVENTIF_EXTENDEDKEY, 0);
                    keybd_event(0x9e, 0x1e, KEYEVENTIF_KEYUP, 0);
                    keybd_event(0x9D, 0x1D, KEYEVENTIF_KEYUP, 0);
                    Thread.Sleep(wait);
                    break;
                case "UP":
                    keybd_event(38, 0, KEYEVENTIF_EXTENDEDKEY, 0);
                    keybd_event(38, 0, KEYEVENTIF_EXTENDEDKEY, 0);
                    Thread.Sleep(wait);
                    break;
                case "DOWN":
                    keybd_event(40, 0, KEYEVENTIF_EXTENDEDKEY, 0);
                    keybd_event(40, 0, KEYEVENTIF_EXTENDEDKEY, 0);
                    Thread.Sleep(wait);
                    break;
                case "RIGHT":
                    keybd_event(39, 0, KEYEVENTIF_EXTENDEDKEY, 0);
                    keybd_event(39, 0, KEYEVENTIF_EXTENDEDKEY, 0);
                    Thread.Sleep(wait);
                    break;
                case "LEFT":
                    keybd_event(37, 0, KEYEVENTIF_EXTENDEDKEY, 0);
                    keybd_event(37, 0, KEYEVENTIF_EXTENDEDKEY, 0);
                    Thread.Sleep(wait);
                    break;
                case "ESCAPE":
                    keybd_event(27, 0, KEYEVENTIF_EXTENDEDKEY, 0);
                    keybd_event(27, 0, KEYEVENTIF_EXTENDEDKEY, 0);
                    Thread.Sleep(wait);
                    break;
            }
        }







    }
}
