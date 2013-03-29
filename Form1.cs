using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Diagnostics;

namespace Remedy_Automation
{
    public partial class Form1 : Form
    {
		//winapi global hotkeys for creating actions. Unregistrated on app close
        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vlc);
        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
		//Find remedy handler for later control.
	    [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindow(string className, string windowTitle);
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindowEx(IntPtr hWnd, IntPtr hwndChild, string className, string windowTitle);
        //Window position detection
		[DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
        //Cursor go to X,Y
		[DllImport("user32.dll")]
        static extern bool SetCursorPos(int X, int Y);
		//Cursor click
		[DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);
		
        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);
		       
        // Hotkeys optional keys
        public const int MOD_ALT = 1;
        public const int MOD_CONTROL = 2;
        public const int MOD_SHIFT = 4;
        public const int MOD_WIN = 8;
		// Hotkeys messages for respond
        public const int WM_HOTKEY1 = 0x312;
        public const int WM_HOTKEY2 = 0x313;
        public const int WM_HOTKEY3 = 0x314;
		// Mouse buttons
        static int WM_LBUTTONDOWN = 0x02;
        static int WM_LBUTTONUP = 0x04;

        public Form1()
        {
            InitializeComponent();
            // Creating global hotkeys after form initialization. 
            RegisterHotKey(this.Handle, WM_HOTKEY1, MOD_ALT, (int)Keys.F7);
            RegisterHotKey(this.Handle, WM_HOTKEY2, MOD_ALT, (int)Keys.F8);
            RegisterHotKey(this.Handle, WM_HOTKEY3, MOD_ALT, (int)Keys.F9);
        }

		//Rectangle structure for window position IN/OUT
        public struct RECT
        {
            public int Left;        // x position of upper-left corner
            public int Top;         // y position of upper-left corner
            public int Right;       // x position of lower-right corner
            public int Bottom;      // y position of lower-right corner
        }

		//Hotkey rections here
        protected override void WndProc(ref Message m)
        {
            //Hotkey check.
            switch ((int)m.WParam)
            {
                case WM_HOTKEY1:
                    IntPtr MainHwnd = FindWindow("ArFrame", "BMC Remedy User - [Система Автоматизированной Эксплуатации Сети  МТС (Поиск)]");
                    IntPtr MakrosHwnd1 = FindWindowEx(MainHwnd, (IntPtr)0, "AfxControlBar70", "");
                    Rectangle myRect = new Rectangle();
                    //IntPtr MakrosHwnd2 = FindWindowEx(MakrosHwnd1, (IntPtr)0, "ToolbarWindow32", "Макрокоманда");
                    //IntPtr Makros_combox = FindWindowEx(MakrosHwnd2, (IntPtr)0, "ComboBox", "");
                    //MessageBox.Show(MakrosHwnd2.ToString());

                    RECT rct;
                    if (!GetWindowRect(MainHwnd, out rct))
                    {
                        MessageBox.Show("Remedy not found. Mb it's not in an acidents screen?");
                        return;
                    }
                    //Отладка отлова позиции окна ремеди
                    //MessageBox.Show("X " + rct.Left.ToString() + "Y " + rct.Top.ToString());
                    myRect.X = rct.Left;
                    myRect.Y = rct.Top;
                    
                    SetCursorPos(myRect.X+500, myRect.Y+70);
                    
                    //Mouse Click to POWER_CREATE
                    mouse_event(WM_LBUTTONDOWN | WM_LBUTTONUP, 0, 0, 0, 0);
                    SetCursorPos(myRect.X + 500, myRect.Y + 85);
                    mouse_event(WM_LBUTTONDOWN | WM_LBUTTONUP, 0, 0, 0, 0);

                    //Кликнуть в элемент сети.
                    //IntPtr MainHwnd = FindWindow("ArFrame", "BMC Remedy User - [Инцидент (Новый)]");
                    SetCursorPos(myRect.X + 482, myRect.Y + 312);
                    mouse_event(WM_LBUTTONDOWN | WM_LBUTTONUP, 0, 0, 0, 0);
                    //Кликнуть в поле ввода
                    SetCursorPos(myRect.X + 469, myRect.Y + 432);
                    mouse_event(WM_LBUTTONDOWN | WM_LBUTTONUP, 0, 0, 0, 0);
                    break;
                case WM_HOTKEY2:
                    label1.Text = Cursor.Position.X.ToString();
                    label2.Text = Cursor.Position.Y.ToString();
                    break;
                case WM_HOTKEY3:
                    MessageBox.Show("F9");
                    break;
            }
            base.WndProc(ref m);
        }

		//Form close
        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            UnregisterHotKey(this.Handle, WM_HOTKEY1);
            UnregisterHotKey(this.Handle, WM_HOTKEY2);
            UnregisterHotKey(this.Handle, WM_HOTKEY3);
        }

        private void notifyIcon1_Click(object sender, EventArgs e)
        {
            this.Show();
            WindowState = FormWindowState.Normal;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized)
            {
                //ShowInTaskbar = false;
                notifyIcon1.BalloonTipText = "Hey, I'm here in tray";
                notifyIcon1.ShowBalloonTip(1000);
            }
            else if (WindowState == FormWindowState.Normal)
            {
                //ShowInTaskbar = true;
            };
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
            //this.Opacity = 0.70;
        }

        private void Form1_Deactivate(object sender, EventArgs e)
        {
            this.Opacity = 0.650;
        }

        private void Form1_Activated(object sender, EventArgs e)
        {
            this.Opacity = 1;
        }
    }
}
