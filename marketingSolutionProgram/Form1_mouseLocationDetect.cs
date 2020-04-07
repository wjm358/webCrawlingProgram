using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace marketingSolutionProgram
{
    partial class Form1
    {

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);

        private const int MOUSEEVENTF_LEFTDOWN = 0x02;
        private const int MOUSEEVENTF_LEFTUP = 0x04;
        private const int MOUSEEVENTF_RIGHTDOWN = 0x08;
        private const int MOUSEEVENTF_RIGHTUP = 0x10;

        System.Timers.Timer mouseDetectTimer = null; //좌표 감지에 쓰이는 타이머


        #region 마우스 클릭 이벤트
        public void LeftDoubleClick(string xpos, string ypos)
        {
            Console.WriteLine("xpos :" + xpos + "  ypos: " + ypos);
            Cursor.Position = new Point(int.Parse(xpos), int.Parse(ypos));

            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            Thread.Sleep(150);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);

            Console.Write("leftmousedoubleclick success");

        }

        public void LeftOneClick(string xpos, string ypos)
        {
            Cursor.Position = new Point(int.Parse(xpos), int.Parse(ypos));
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);

        }
        #endregion


        #region 마우스 좌표탐지
        //실시간 좌표 감지 시작
        private void mouseDetectStart(object sender, EventArgs e)
        {
            if(mouseDetectTimer == null )
            {
            mouseDetectTimer = setTimer();
            mouseDetectTimer.Elapsed += timer_Elapsed;
            mouseDetectTimer.Start();
            }
        }
        //실시간 좌표 감지 종료
        private void mouseDetectStop(object sender, EventArgs e)
        {
            if(mouseDetectTimer != null)
            {
            mouseDetectTimer.Stop();
            mouseDetectTimer.Dispose();
            mouseDetectTimer = null;
            }
        }
        delegate void TimerEventFiredDelegate();
        void timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            BeginInvoke(new TimerEventFiredDelegate(Work));
        }
        private void Work()
        {
            xAbsLocLabel.Text = "X=" + Cursor.Position.X.ToString();
            yAbsLocLabel.Text = "Y= " + Cursor.Position.Y.ToString();
            //수행해야할 작업(UI Thread 핸들링 가능)
        }
        #endregion
    }
}
