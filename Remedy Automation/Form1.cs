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

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool Beep(int frequency, int duration);
		       
        // Hotkeys optional keys
        public const int MOD_ALT = 1;
        public const int MOD_CONTROL = 2;
        public const int MOD_SHIFT = 4;
        public const int MOD_WIN = 8;
		// Hotkeys messages for respond
        public const int WM_HOTKEY1 = 0x312;
        public const int WM_HOTKEY2 = 0x313;
        public const int WM_HOTKEY3 = 0x314;
        public const int WM_HOTKEY4 = 0x315;
        public const int WM_HOTKEY5 = 0x316;
		// Mouse buttons
        static int WM_LBUTTONDOWN = 0x02;
        static int WM_LBUTTONUP = 0x04;
        
        bool Expand_flag = true;

        public Form1()
        {
            InitializeComponent();
            // Creating global hotkeys after form initialization. 
            RegisterHotKey(this.Handle, WM_HOTKEY1, 0, (int)Keys.F8);
            RegisterHotKey(this.Handle, WM_HOTKEY2, 0, (int)Keys.F7);
            RegisterHotKey(this.Handle, WM_HOTKEY3, MOD_ALT, (int)Keys.F8);
            RegisterHotKey(this.Handle, WM_HOTKEY4, 0, (int)Keys.F9);
            RegisterHotKey(this.Handle, WM_HOTKEY5, 0, (int)Keys.F10);
        }

		//Rectangle structure for window position IN/OUT
        public struct RECT
        {
            public int Left;        // x position of upper-left corner
            public int Top;         // y position of upper-left corner
            public int Right;       // x position of lower-right corner
            public int Bottom;      // y position of lower-right corner
        }

        static void Click(int x, int y)
        {
            SetCursorPos(x, y);
            mouse_event(WM_LBUTTONDOWN | WM_LBUTTONUP, 0, 0, 0, 0);
        }

        static void DoubleClick(int x, int y)
        {
            SetCursorPos(x, y);
            mouse_event(WM_LBUTTONDOWN | WM_LBUTTONUP, 0, 0, 0, 0);
            mouse_event(WM_LBUTTONDOWN | WM_LBUTTONUP, 0, 0, 0, 0);
        }
        
        Rectangle myRect = new Rectangle();
        RECT rct;

		//Hotkey rections here
        protected override void WndProc(ref Message m)
        {
            //Hotkey check.
            switch ((int)m.WParam)
            {
                case WM_HOTKEY2:

                    IntPtr MainHwnd = FindWindow("ArFrame", "BMC Remedy User - [Система Автоматизированной Эксплуатации Сети  МТС (Поиск)]");
                    if (!GetWindowRect(MainHwnd, out rct))
                    {
                        MessageBox.Show(MainHwnd.ToString());
                        return;
                    }

                    myRect.X = rct.Left;
                    myRect.Y = rct.Top;

                    //Создать по шаблону
                    Click(myRect.X+50, myRect.Y+305);
					//Откючение питания
                    DoubleClick(myRect.X + 137, myRect.Y + 245);
                    //Кликнуть в элемент сети.
                    Click(myRect.X + 482, myRect.Y + 312);
                    //Кликнуть в поле ввода
                    Click(myRect.X + 469, myRect.Y + 432);
                    break;
                case WM_HOTKEY1:
                    MainHwnd = FindWindow("ArFrame", "BMC Remedy User - [Инцидент (Новый)]");
                    if (!GetWindowRect(MainHwnd, out rct))
                    {
						//%Username%, ты делаешь что-то не так
                        MessageBox.Show("Remedy not found. Mb it's not in an acidents screen?");
                        return;
                    }
                    myRect.X = rct.Left;
                    myRect.Y = rct.Top;
                    //Прокликиваем какой-то шлак
                    Click(myRect.X + 2344-1680, myRect.Y + 621);
                    //Выделение  Радиоподсистема -> Прочие. 
                    Click(myRect.X + 228, myRect.Y + 695);
					InputLanguage.CurrentInputLanguage =  InputLanguage.FromCulture(new System.Globalization.CultureInfo("ru-RU"));
                    System.Threading.Thread.Sleep(500);
                    SendKeys.Send("Прочие");
                    //3
                    Click(myRect.X + 2462-1680, myRect.Y + 625);
                    //4
                    Click(myRect.X + 2179-1680, myRect.Y + 746);
                    //5
                    Click(myRect.X + 2174-1680, myRect.Y + 742);
                    //изменить масштаб влияния на 1
                    //6
                    Click(myRect.X + 314, myRect.Y + 901);
                    //7
                    Click(myRect.X + 314, myRect.Y + 911);
                    break;
                case WM_HOTKEY3:
                    label1.Text = "X = " + Cursor.Position.X.ToString();
                    label2.Text = "Y = " + Cursor.Position.Y.ToString();
                    break;
                case WM_HOTKEY4:
                    MainHwnd = FindWindow("ArFrame", "BMC Remedy User - [Система Автоматизированной Эксплуатации Сети  МТС (Поиск)]");
                    if (!GetWindowRect(MainHwnd, out rct))
                    {
                        MessageBox.Show(MainHwnd.ToString());
                        return;
                    }
                    
                    myRect.X = rct.Left;
                    myRect.Y = rct.Top;

                    //Mouse Move to create incident link
                    //1
                    Click(myRect.X+50, myRect.Y+305);
                    //2
                    DoubleClick(myRect.X + 137, myRect.Y + 544);
                    //3
                    Click(myRect.X + 463, myRect.Y + 402);
                    //4
                    Click(myRect.X + 345, myRect.Y + 795);
                    //5
                    Click(myRect.X + 493, myRect.Y + 557);
                    //Кликнуть в элемент сети.
                    //IntPtr MainHwnd = FindWindow("ArFrame", "BMC Remedy User - [Инцидент (Новый)]");
                    //6
                    Click(myRect.X + 482, myRect.Y + 312);
                    //Кликнуть в поле ввода
                    //7
                    Click(myRect.X + 469, myRect.Y + 432);
                    break;
                case WM_HOTKEY5:
                    //1
                    Click(myRect.X + 2462-1680, myRect.Y + 625);
                    //2
                    Click(myRect.X + 314, myRect.Y + 901);
                    //3
                    Click(myRect.X + 314, myRect.Y + 911);
                    break;
            }
            base.WndProc(ref m);
        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            UnregisterHotKey(this.Handle, WM_HOTKEY1);
            UnregisterHotKey(this.Handle, WM_HOTKEY2);
            UnregisterHotKey(this.Handle, WM_HOTKEY3);
            notifyIcon1.Visible = false;
        }

        private void notifyIcon1_Click(object sender, EventArgs e)
        {
            this.Show();
            WindowState = FormWindowState.Normal;
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

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (Expand_flag)
            {
                this.Width = 700;
                Expand_flag = false;
            }
            else
            {
                this.Width = 245;
                Expand_flag = true;
            }

        }
    }
}
