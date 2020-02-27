using mshtml;
using SHDocVw;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using HtmlAgilityPack;
using System.Runtime.InteropServices;
using System.IO;
using System.Net;
using System.Net.Sockets;
using log4net;
using System.Text.RegularExpressions;
using System.Security.Principal;

[assembly: log4net.Config.XmlConfigurator(Watch = true)]

namespace marketingSolutionProgram
{
    public partial class Form1 : Form
    {
        private ILog log = LogManager.GetLogger("Program");
        BackgroundWorker worker;

        #region win32api 선언
        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, int wParam, int lParam);
        private const int BM_CLICK = 0xF5;
        private const uint WM_ACTIVATE = 0x6;
        private const int WA_ACTIVE = 1;



        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern IntPtr SetFocus(IntPtr handle);
        [DllImport("user32.dll")]
        static extern IntPtr GetFocus();
        [DllImport("user32.dll")]
        private static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

        private const int SW_SHOWNORMAL = 1;
        private const int SW_SHOWMINIMIZED = 2;
        private const int SW_SHOWMAXIMIZED = 3;

        [Flags]
        public enum KeyFlag
        {
            /// <summary>
            /// 키 누름
            /// </summary>
            KE_DOWN = 0,
            /// <summary>
            /// 확장 키
            /// </summary>
            KE__EXTENDEDKEY = 1,
            /// <summary>
            /// 키 뗌
            /// </summary>
            KE_UP = 2
        }

        [DllImport("User32.dll")]
        static extern void keybd_event(byte vk, byte scan, int flags, int extra);

        const int WM_GETTEXT = 0x000D;
        const int WM_GETTEXTLENGTH = 0x000E;
        const int WM_SETTEXT = 0x000C;

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);

        private const int MOUSEEVENTF_LEFTDOWN = 0x02;
        private const int MOUSEEVENTF_LEFTUP = 0x04;
        private const int MOUSEEVENTF_RIGHTDOWN = 0x08;
        private const int MOUSEEVENTF_RIGHTUP = 0x10;
        #endregion

        int indexNum = 0;
        int macroListTextboxCursor = 0;

        string indexString = String.Empty;
        string checkString = String.Empty;
        string macroString = String.Empty;
        string macroList = String.Empty;
        string prevUrl = String.Empty;

        private Stopwatch totalsw = null;
        Microsoft.Win32.RegistryKey rk = null;
        System.Timers.Timer mouseDetectTimer = null; //좌표 감지에 쓰이는 타이머
        Random random = null;
        Thread workerThread = null;
        Thread aa = null;

        InternetExplorer ie = null;
        SHDocVw.WebBrowser webBrowser = null;
        HtmlAgilityPack.HtmlDocument document;
        HtmlNodeCollection nodes;
        //https://m.comic.naver.com/external/appLaunchBridge.nhn?type=ARTICLE_LIST&titleId=651073
        #region 변수 재활용
        private Stopwatch setTotalSw()
        {
            if (totalsw == null)
                totalsw = new Stopwatch();
            totalsw.Reset();
            return totalsw;
        }
        private Random setRandomInstance()
        {
            if (random == null)
                random = new Random(Guid.NewGuid().GetHashCode());
            return random;
        }
        private System.Timers.Timer setTimer()
        {
            if (mouseDetectTimer == null)
                mouseDetectTimer = new System.Timers.Timer();
            return mouseDetectTimer;
        }
        #endregion

        #region 체류시간 설정
        private void sleep(int s, int e)
        {
            random = setRandomInstance();
            int randomSecond = random.Next(s, e);
            Thread.Sleep(randomSecond * 1000);
            return;
        }
        #endregion

        #region win32api 마우스 메소드
        string ReplacementHangulFragment(char Character)
        {
            //            static const TCHAR Chosung[19][3]={{L"r"}, {L"R"}, {L"s"}, {L"e"}, {L"E"}, {L"f"}, {L"a"}, {L"q"}, {L"Q"}, {L"t"}, {L"T"}, {L"d"}, {L"w"}, {L"W"}, {L"c"}, {L"z"}, {L"x"}, {L"v"}, {L"g"}};
            //	static const TCHAR Joongsung[21][3]={{L"k"}, {L"o"}, {L"i"}, {L"O"}, {L"j"}, {L"p"}, {L"u"}, {L"P"}, {L"h"}, {L"hk"}, {L"ho"}, {L"hl"}, {L"y"}, {L"n"}, {L"nj"}, {L"np"}, {L"nl"}, {L"b"}, {L"m"}, {L"ml"}, {L"l"}};
            //	static const TCHAR Jongsung[28][3]={{L" "}, {L"r"}, {L"R"}, {L"rt"}, {L"s"}, {L"sw"}, {L"sg"}, {L"e"}, {L"f"}, {L"fr"}, {L"fa"}, {L"fq"}, {L"ft"}, {L"fx"}, {L"fv"}, {L"fg"}, {L"a"}, {L"q"}, {L"qt"}, {L"t"}, {L"T"}, {L"d"}, {L"w"}, {L"c"}, {L"z"}, {L"x"}, {L"v"}, {L"g"}};
            //	const int CommonNumber = Character - 0xAC00;
            //static TCHAR ReplacementFragment[10];

            //int ChosungIndex = (int)(CommonNumber / (28 * 21));
            //int JoongsungIndex = (int)((CommonNumber % (28 * 21)) / 28);
            //int JongsungIndex = (int)(CommonNumber % 28);

            //wcscpy(ReplacementFragment, Chosung[ChosungIndex]);
            //wcscat(ReplacementFragment, Joongsung[JoongsungIndex]);
            //	if(JongsungIndex != 0 )
            //	{
            //		wcscat(ReplacementFragment, Jongsung[JongsungIndex]);
            //	}

            //	return ReplacementFragment;
            return string.Empty;
        }

        void TypingEnglish(string Message)
        {
            string NumSpecialChar = ")!@#$%^&*(";
            string SpecialChar = "`-=\\[];',./";
            string ShiftSpecialChar = "~_+|{}:\"<>?";
            byte[] SpecialCharCode = new byte[] { 0xC0, 0xBD, 0xBB, 0xDC, 0xDB, 0xDD, 0xBA, 0xDE, 0xBC, 0xBE, 0xBF };

            int length = Message.Length;
            for (int i = 0; i < length; i++)
            {
                if (Message[i] >= 'a' && Message[i] <= 'z')
                {
                    keybd_event((byte)(Message[i] - 'a' + 'A'), 0, 0, 0);
                    keybd_event((byte)(Message[i] - 'a' + 'A'), 0, (int)KeyFlag.KE_UP, 0);
                }
            }
            //          int Length = wcslen(Message);

            //          for (int i = 0; i < Length; i++)
            //          {
            //              if (Message[i] >= L'a' && Message[i] <= L'z' )
            //{
            //              keybd_event((BYTE)Message[i] - L'a' + L'A', 0, 0, 0);
            //              keybd_event((BYTE)Message[i] - L'a' + L'A', 0, KEYEVENTF_KEYUP, 0);
            //          }
            //else if (Message[i] >= L'A' && Message[i] <= L'Z' )
            //{
            //              keybd_event(VK_SHIFT, 0, 0, 0);
            //              keybd_event((BYTE)Message[i], 0, 0, 0);
            //              keybd_event((BYTE)Message[i], 0, KEYEVENTF_KEYUP, 0);
            //              keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0);
            //          }
            //else if (Message[i] >= L'0' && Message[i] <= L'9' )
            //{
            //              keybd_event((BYTE)Message[i], 0, 0, 0);
            //              keybd_event((BYTE)Message[i], 0, KEYEVENTF_KEYUP, 0);
            //          }
            //else if (Message[i] == L' ' )
            //{
            //              keybd_event((BYTE)Message[i], 0, 0, 0);
            //              keybd_event((BYTE)Message[i], 0, KEYEVENTF_KEYUP, 0);
            //          }
            //else
            //{
            //              for (int j = 0; j < 10; j++)
            //              {
            //                  if (Message[i] == NumSpecialChar[j])
            //                  {
            //                      keybd_event(VK_SHIFT, 0, 0, 0);
            //                      keybd_event(L'0' + j, 0, 0, 0);
            //                      keybd_event(L'0' + j, 0, KEYEVENTF_KEYUP, 0);
            //                      keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0);
            //                      break;
            //                  }
            //                  if (j == 9)
            //                  {
            //                      for (int k = 0; k < 11; k++)
            //                      {
            //                          if (Message[i] == SpecialChar[k])
            //                          {
            //                              keybd_event(SpecialCharCode[k], 0, 0, 0);
            //                              keybd_event(SpecialCharCode[k], 0, KEYEVENTF_KEYUP, 0);
            //                              break;
            //                          }
            //                          else if (Message[i] == ShiftSpecialChar[k])
            //                          {
            //                              keybd_event(VK_SHIFT, 0, 0, 0);
            //                              keybd_event(SpecialCharCode[k], 0, 0, 0);
            //                              keybd_event(SpecialCharCode[k], 0, KEYEVENTF_KEYUP, 0);
            //                              keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0);
            //                              break;
            //                          }
            //                      }
            //                  }
            //              }
            //          }
            //}
        }

        void TypingMessage(string Message)
        {
            //static TCHAR Buf[10];
            //int Length = wcslen(Message);

            //for (int i = 0; i < Length; i++)
            //{
            //    if (Message[i] >= 0xAC00 && Message[i] <= 0xD7A3)
            //    {
            //        wcscpy(Buf, ReplacementHangulFragment(Message[i]));
            //        TypingEnglish(Buf);
            //    }
            //    else
            //    {
            //        Buf[0] = Message[i];
            //        Buf[1] = L'\0';
            //        keybd_event(VK_HANGEUL, MapVirtualKey(VK_HANGEUL, 0), 0, 0);
            //        keybd_event(VK_HANGEUL, MapVirtualKey(VK_HANGEUL, 0), KEYEVENTF_KEYUP, 0);
            //        TypingEnglish(Buf);
            //        keybd_event(VK_HANGEUL, MapVirtualKey(VK_HANGEUL, 0), 0, 0);
            //        keybd_event(VK_HANGEUL, MapVirtualKey(VK_HANGEUL, 0), KEYEVENTF_KEYUP, 0);
            //    }
            //}
        }

        #endregion

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

        #region 백그라운드 스레드 정의
        /// <summary>
        /// Delegate for the EnumChildWindows method
        /// </summary>
        /// <param name="hWnd">Window handle</param>
        /// <param name="parameter">Caller-defined variable; we use it for a pointer to our list</param>
        /// <returns>True to continue enumerating, false to bail.</returns>
        public delegate bool EnumWindowProc(IntPtr hWnd, IntPtr parameter);
        [DllImport("user32")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumChildWindows(IntPtr window, EnumWindowProc callback, IntPtr i);

        /// <summary>
        /// Returns a list of child windows
        /// </summary>
        /// <param name="parent">Parent of the windows to return</param>
        /// <returns>List of child windows</returns>
        public static List<IntPtr> GetChildWindows(IntPtr parent)
        {
            List<IntPtr> result = new List<IntPtr>();
            GCHandle listHandle = GCHandle.Alloc(result);
            try
            {
                EnumWindowProc childProc = new EnumWindowProc(EnumWindow);
                EnumChildWindows(parent, childProc, GCHandle.ToIntPtr(listHandle));
            }
            finally
            {
                if (listHandle.IsAllocated)
                    listHandle.Free();
            }
            return result;
        }
        /// <summary>
        /// Callback method to be used when enumerating windows.
        /// </summary>
        /// <param name="handle">Handle of the next window</param>
        /// <param name="pointer">Pointer to a GCHandle that holds a reference to the list to fill</param>
        /// <returns>True to continue the enumeration, false to bail</returns>
        private static bool EnumWindow(IntPtr handle, IntPtr pointer)
        {
            GCHandle gch = GCHandle.FromIntPtr(pointer);
            List<IntPtr> list = gch.Target as List<IntPtr>;
            if (list == null)
            {
                throw new InvalidCastException("GCHandle Target could not be cast as List<IntPtr>");
            }
            list.Add(handle);
            //  You can modify this to check to see if you want to cancel the operation, then return a null here
            return true;
        }


        [DllImport("user32")]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
        [DllImport("user32")]
        public static extern IntPtr GetWindow(IntPtr hwnd, int wCmd);
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        StringBuilder sb = new StringBuilder(1010);
        StringBuilder tempSb = new StringBuilder(1010);
        StringBuilder IEFrameWindowName = new StringBuilder(1010);
        void deleteAlert()
        {
            sb.Clear();
            tempSb.Clear();
            IEFrameWindowName.Clear();
            //바탕화면 핸들 
            IntPtr mainHwnd = FindWindow("#32769", null);
            List<IntPtr> allList = GetChildWindows(mainHwnd);

            List<IntPtr> IEFrameLists = new List<IntPtr>();
            List<IntPtr> alertLists = new List<IntPtr>();

            //IEFrame 찾아서 저장
            foreach (var tempHwnd in allList)
            {
                tempSb.Clear();
                GetClassName(tempHwnd, tempSb, 1010);
                if (tempSb.ToString().CompareTo("IEFrame") == 0)
                {
                    IEFrameLists.Add(tempHwnd);
                    continue;
                }
                if (tempSb.ToString().CompareTo("#32770") == 0)
                {
                    alertLists.Add(tempHwnd);
                    continue;
                }
            }

            //  Console.WriteLine(IEFrameLists.Count + " , " + alertLists.Count);


            foreach (var IEFrameHwnd in IEFrameLists)
            {
                IEFrameWindowName.Clear();
                GetWindowText(IEFrameHwnd, IEFrameWindowName, 1010);
                //    Console.WriteLine(IEFrameWindowName);
                foreach (var alertHwnd in alertLists)
                {
                    tempSb.Clear();
                    var ownerHwnd = GetWindow(alertHwnd, 4);
                    if (ownerHwnd != IntPtr.Zero)
                    {
                        GetWindowText(ownerHwnd, tempSb, 1010);
                        if (tempSb.ToString().CompareTo(IEFrameWindowName.ToString()) == 0)
                        {
                            SetForegroundWindow(alertHwnd);
                            var buttonHwnd = FindWindowEx(alertHwnd, IntPtr.Zero, "Button", null);
                            if (buttonHwnd != IntPtr.Zero)
                            {
                                //Console.WriteLine("not zero");
                            }
                            SendMessage(buttonHwnd, WM_ACTIVATE, WA_ACTIVE, 0);

                            SendMessage(buttonHwnd, BM_CLICK, 0, 0);


                        }
                    }
                }
            }

        }

        #endregion

        public Form1()
        {
            InitializeComponent();

            deleteAllIeProcesses();
        }

        //ie프로세스 모두 삭제
        public void deleteAllIeProcesses()
        {
            Process[] IeProcesses = Process.GetProcessesByName("iexplore");
            foreach (var IeProcess in IeProcesses)
            {
                IeProcess.Kill();

            }
        }

        //check iframe length
        public bool isPresentIframe()
        {
            int count = 0;

         
                document = new HtmlWeb().Load(ie.LocationURL);
                nodes = document.DocumentNode.SelectNodes("//iframe[@src]");

                if (nodes != null)
                    count = nodes.Count;

        

            return count > 0;
        }

        //iframe 갯수 
        public int getIframeCount()
        {
            int count = 0;
            document = new HtmlWeb().Load(webBrowser.LocationURL);
            nodes = document.DocumentNode.SelectNodes("//iframe[@src]");
            count = nodes.Count;

            return count;

        }

        #region handler thread
        public bool ExplorerMessage()
        {
            var hwnd = FindWindow("#32770", "Internet Explorer");
            if (hwnd != IntPtr.Zero)
            {

                SetForegroundWindow((IntPtr)ie.HWND);
                SendKeys.SendWait("{ESC}");
                SendKeys.SendWait("{ESC}");
                SendKeys.SendWait("{ESC}");
                // find button on dialog window: classname = "Button", text = "OK"
                var btn = FindWindowEx(hwnd, IntPtr.Zero, "Button", "취소");

                if (btn != IntPtr.Zero)
                {

                    // activate the button on dialog first or it may not acknowledge a click msg on first try
                    SendMessage(btn, WM_ACTIVATE, WA_ACTIVE, 0);
                    // send button a click message

                    SendMessage(btn, BM_CLICK, 0, 0);
                    return true;
                }
                return true;
            }
            return false;

        }
        public void checkHttpError()
        {
            if (ie != null)
            {
                string title = ie.LocationName;
                if (title != null)
                {
                    if (title.Contains("HTTP Status"))
                    {
                        makeIeProcess(prevUrl);
                        log.Debug("HTTP Error 처리 완료");
                    }
                }
            }
        }
        //close alert when pop up alert
        public void ActivateAndClickOkButton()
        {
            // find dialog window with titlebar text of "Message from webpage"

            for (int i = 0; i < 6; i++)
            {
                checkHttpError();

                if (ExplorerMessage() == true)
                {
                    Console.WriteLine("explorerMessage true");
                    log.Debug("exlorerMessage true");
                    return;
                }

                var hwnd = FindWindow("#32770", "웹 페이지 메시지");
                if (hwnd != IntPtr.Zero)
                {
                    SetForegroundWindow((IntPtr)ie.HWND);
                    SendKeys.SendWait("{ESC}");

                    Console.WriteLine("esc complete");
                    log.Debug("esc complete");
                    return;

                    //// find button on dialog window: classname = "Button", text = "OK"
                    //var btn = FindWindowEx(hwnd, IntPtr.Zero, "Button", "취소");


                    //if (btn != IntPtr.Zero)
                    //{

                    //    // activate the button on dialog first or it may not acknowledge a click msg on first try
                    //    SendMessage(btn, WM_ACTIVATE, WA_ACTIVE, 0);
                    //    // send button a click message

                    //    SendMessage(btn, BM_CLICK, 0, 0);
                    //    return;
                    //}
                    //else
                    //{
                    //    btn = FindWindowEx(hwnd, IntPtr.Zero, "Button", "확인");
                    //    if (btn != IntPtr.Zero)
                    //    {
                    //        // activate the button on dialog first or it may not acknowledge a click msg on first try
                    //        SendMessage(btn, WM_ACTIVATE, WA_ACTIVE, 0);
                    //        // send button a click message

                    //        SendMessage(btn, BM_CLICK, 0, 0);
                    //        return;
                    //    }
                    //    //btn = FindWindowEx(hwnd,IntPtr.Zero, "Button", )
                    //    //Interaction.MsgBox("button not found!");
                    //}
                }
                else
                {
                    Thread.Sleep(1000);

                    //Interaction.MsgBox("window not found!");
                }
            }

        }
        #endregion

        bool findByKeyword(string keyword)
        {
            HtmlWeb web = new HtmlWeb();
            mshtml.HTMLDocument aa = (mshtml.HTMLDocument)webBrowser.Document;

            var documentAsIHtmlDocument3 = (mshtml.IHTMLDocument3)aa;
            StringReader sr = new StringReader(documentAsIHtmlDocument3.documentElement.outerHTML);
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.Load(sr);

            Console.WriteLine("findByKeyword :" + ie.LocationURL);
            //doc = web.Load(ie.LocationURL);
            //doc.LoadHtml(ie.Document.documentElement.outerHTML);
            //HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            //doc.LoadHtml(ie.Document.DocumentElement.outerHTML);

            var keywordElement = doc.DocumentNode.SelectSingleNode("//*[text()[contains(., '" + keyword + "' )]]");

            string tagName = null;
            //mindocument에 없는 것임
            if (keywordElement == null)

            {
                Console.WriteLine("keywordElement is null");
                //doesn't exist in mainDocument
                if (isPresentIframe() == true)
                {
                    Console.WriteLine("isPresentIframe method returned true");

                    mshtml.HTMLDocument doc1 = (mshtml.HTMLDocument)ie.Document;
                    FramesCollection framesList = doc1.frames;

                    for (int i = 0; i < framesList.length; i++)
                    {
                        Console.WriteLine(i);
                        object index = (object)i;

                        mshtml.HTMLWindow2 frame = (mshtml.HTMLWindow2)framesList.item(ref index);
                        doc1 = (mshtml.HTMLDocument)frame.document;


                        var htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        htmlDoc.LoadHtml(doc1.documentElement.outerHTML);

                        keywordElement = htmlDoc.DocumentNode.SelectSingleNode("//*[text()[contains(., '" + keyword + "' )]]");
                        if (keywordElement != null) tagName = keywordElement.Name;

                        if (tagName == null) continue;

                        var elements = doc1.getElementsByTagName(tagName);

                        foreach (IHTMLElement element in elements)
                        {
                            string href = element.innerText;

                            if (href != null && keywordElement.InnerText != null && href.CompareTo(keywordElement.InnerText.Replace("\r\n", "")) == 0)
                            {
                                element.click();

                                return true;
                            }
                        }
                    }
                }


            }
            else
            {
                Console.WriteLine("keywordElement is not null");
                //exist in mainDocument
                tagName = keywordElement.Name;

                var elements = aa.getElementsByTagName(tagName);
                Console.WriteLine(tagName + "," + elements.length);
                //main document에서 조사

                foreach (IHTMLElement temp in elements)
                {
                    string tempString = temp.innerText;
                    if (tempString != null)
                    {
                        tempString = tempString.Trim();

                    }
                    //string tempString = temp.innerText.Replace("\r\n","").Trim();

                    //if(temp.innerText !=null)
                    //{
                    //    if(temp.innerText.CompareTo("여행맛집")== 0)
                    //    {
                    //        Console.WriteLine("correct");
                    //    }

                    //}
                    if (tempString != null && tempString.CompareTo(keyword) == 0)
                    {

                        Console.WriteLine("clicked");
                        temp.click();


                        return true;
                    }
                }


            }



            return false;
        }

        public void findByHref(string keyword)
        {
            mshtml.HTMLDocument doc = (mshtml.HTMLDocument)ie.Document;
            var elements = doc.getElementsByTagName("a");

            keyword = makeToHrefString(ref keyword);
            string hrefString = string.Empty;

            foreach (IHTMLElement temp in elements)
            {
                hrefString = temp.getAttribute("href");

                if (hrefString != null && hrefString.CompareTo(keyword) == 0)
                {
                    temp.click();

                    return;
                }
            }

            //main document에 존재하지 않음
            if (isPresentIframe() == true)
            {
                mshtml.HTMLDocument doc1 = (mshtml.HTMLDocument)ie.Document;
                FramesCollection framesList = doc1.frames;

                for (int i = 0; i < framesList.length; i++)
                {
                    object index = i;
                    mshtml.IHTMLWindow2 frame = (mshtml.IHTMLWindow2)doc1.frames.item(ref index);
                    doc1 = (mshtml.HTMLDocument)frame.document;

                    var aTags = doc1.getElementsByTagName("a");

                    foreach (IHTMLElement a in aTags)
                    {
                        string href = a.getAttribute("href");

                        if (href != null && href.CompareTo(keyword) == 0)
                        {
                            a.click();

                            return;
                        }
                    }
                }
            }
        }

        //href format 체크 후 convert
        public string makeToHrefString(ref string keyword)
        {
            if (keyword.StartsWith("href"))
            {
                keyword = keyword.Substring(keyword.IndexOf("="));
            }
            return keyword;

        }

        public int getRandomIndex(int elementsLength)
        {
            random = setRandomInstance();
            int randomIndex = random.Next(0, elementsLength);
            return randomIndex;
        }

        public bool elementClick(int cnt, int randomIndex, IHTMLElementCollection elements)
        {
            int count = 0;
            foreach (mshtml.IHTMLDOMNode element in elements)
            {
                if (count == randomIndex)
                {
                    //IHTMLAttributeCollection attrs = element.attributes;
                    //if (attrs != null)
                    //{
                    //    foreach (IHTMLDOMAttribute at in attrs)
                    //    {
                    //        if (at.specified)
                    //        {

                    //            string nodeValue = "";
                    //            if (at.nodeValue != null)
                    //                nodeValue = at.nodeName.ToString();
                    //            if (nodeValue.CompareTo("onclick") == 0)
                    //            {
                    //                Console.WriteLine(cnt + ", onclick");
                    //                return false;
                    //            }
                    //        }
                    //    }
                    //}

                    ((mshtml.IHTMLElement)element).click();
                    return true;

                }
                count++;
            }

            return false;
        }

        #region 테스트 클릭 버튼 메소드
        public void elementClickRepeatA()
        {
            mshtml.HTMLDocument doc = (mshtml.HTMLDocument)ie.Document;
            bool result = false;
            int count = 0;
            while (result == false && count < 10)
            {

                var elements = doc.getElementsByTagName("a");
                int randomIndex = getRandomIndex(elements.length);

                result = elementClick(count, randomIndex, elements);
                count++;

            }


        }
        public void elementClickRepeatB()
        {
            mshtml.HTMLDocument doc = (mshtml.HTMLDocument)ie.Document;
            bool result = false;
            int count = 0;
            //iframe에서 수행
            int iframeCount = getIframeCount();
            int randomIndex = getRandomIndex(iframeCount);
            object index = (object)randomIndex;
            mshtml.IHTMLWindow2 frame = (mshtml.IHTMLWindow2)doc.frames.item(ref index);
            while (result == false && count < 10)
            {


                if (frame.document != null)
                    doc = (mshtml.HTMLDocument)frame.document;
                else return;

                var aTags = doc.getElementsByTagName("a");
                randomIndex = getRandomIndex(aTags.length);
                result = elementClick(count, randomIndex, aTags);
                count++;
            }


        }
        public void clickRandomlyExceptOnclick()
        {
            if (isPresentIframe() == false)
            {
                Console.WriteLine("isPresentIFrame false");
                elementClickRepeatA();

            }
            else
            {
                //iframe이 존재
                //http://rpp.gmarket.co.kr/?exhib=33275 <<- about:blank 제외해야함
                int randomIndex = getRandomIndex(2);
                bool mainDocumentIsChoiced = (randomIndex == 0 ? false : true);

                if (mainDocumentIsChoiced == true)
                {

                    elementClickRepeatA();
                }
                else
                {
                    // https://m.sports.naver.com/news.nhn?oid=139&aid=0002128129
                    elementClickRepeatB();
                }

            }
        }
        #endregion

        public void clickRandomly()
        {

            if (isPresentIframe() == false)
            {
                //main document 에서 실행 <- iframe이 존재하지 않으므로
                mshtml.HTMLDocument doc = (mshtml.HTMLDocument)ie.Document;
                var elements = doc.getElementsByTagName("a");
                int randomIndex = getRandomIndex(elements.length);
                elementClick(0, randomIndex, elements);

            }
            else
            {
                //iframe이 존재
                //http://rpp.gmarket.co.kr/?exhib=33275 <<- about:blank 제외해야함
                int randomIndex = getRandomIndex(2);
                bool mainDocumentIsChoiced = (randomIndex == 0 ? false : true);

                if (mainDocumentIsChoiced == true)
                {

                    //maindocument에서 수행 
                    mshtml.HTMLDocument doc = (mshtml.HTMLDocument)ie.Document;
                    var elements = doc.getElementsByTagName("a");
                    randomIndex = getRandomIndex(elements.length);
                    elementClick(0, randomIndex, elements);
                }
                else
                {
                    // https://m.sports.naver.com/news.nhn?oid=139&aid=0002128129

                    //iframe에서 수행
                    int iframeCount = getIframeCount();
                    randomIndex = getRandomIndex(iframeCount);
                    object index = (object)randomIndex;

                    mshtml.HTMLDocument doc = (mshtml.HTMLDocument)ie.Document;
                    mshtml.IHTMLWindow2 frame = (mshtml.IHTMLWindow2)doc.frames.item(ref index);


                    doc = (mshtml.HTMLDocument)frame.document;


                    var aTags = doc.getElementsByTagName("a");
                    randomIndex = getRandomIndex(aTags.length);
                    elementClick(0, randomIndex, aTags);
                }

            }

        }

        //캐시 삭제
        public void clearHistory()
        {
            deleteAllTab();
            ie = null;

            System.Diagnostics.Process.Start("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 4351").WaitForExit();
            Console.WriteLine("history clear");
            log.Debug("캐시 삭제 완료");
        }

        private string checkUrl(string macroString)
        {

            bool includeHttps = macroString.StartsWith("http");
            if (includeHttps == false)
            {
                bool includeHttp = macroString.StartsWith("http");

                return "http://" + macroString;
            }
            else
            {
                return macroString;
            }
        }

        #region BackgroundWorker event define

        bool checkRegex(string checkString, string pattern)
        {
            return Regex.IsMatch(checkString, pattern);
        }

        void bw_DoWork(List<string> splitMacroList, int macroListLength, DoWorkEventArgs e)
        {
            label11.Text = "총 명령어 개수 : \n";
            label11.Text += macroListLength;
            rk =
          Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Internet Settings\5.0\User Agent", true);
            if (radioButton2.Checked == true)
            {
                //"Mozilla/5.0 (Linux; Android 10.0.0; SGH-i907) AppleWebKit/664.76 (KHTML, like Gecko) Chrome/87.0.3131.15 Mobile Safari/664.76 
                //"Mozilla /5.0(Linux; U; Android 7.0.0; SGH-i907) AppleWebKit / 533.1(KHTML, like Gecko) Version / 4.0 Mobile Safari/ 533.1 Chrome/87.0.3131.15"
                rk.SetValue(null, "Mozilla/5.0 (Linux; Android 10.0.0; SGH-i907) AppleWebKit/664.76 (KHTML, like Gecko) Chrome/87.0.3131.15 Mobile Safari/664.76");
                if (rk != null)
                {
                    log.Debug("레지스트리 수정 완료");

                };
            }
            else
                rk.SetValue(null, "");

                aa = new Thread(() =>
               {
                   while (true)
                   {
                       try
                       {

                           deleteAlert();
                       }
                       catch (Exception ex)
                       {
                           log.Debug(ex);
                       }

                       Thread.Sleep(2000);
                   }
               });
            aa.IsBackground = true;
            aa.Start();

            while (true)

            {
                totalsw = setTotalSw();
                totalsw.Start();

                log.Debug("\n\n\n매크로 작업목록 출력");
                foreach (string temp in splitMacroList)
                {
                    log.Debug(temp);
                }
                log.Debug("----------------------------------------------------\n\n");

                for (int i = 0; i < macroListLength; i++)
                {
                    //취소버튼 클릭시
                    if (worker.CancellationPending)
                    {
                        e.Cancel = true;
                        totalsw.Stop();
                        return;
                    }


                    macroString = splitMacroList[i].Trim().Substring(1); //특수문자 제거
                    printCurrentMacro(macroString, i); //현재 명령을 라벨에 출력

                    indexNum = macroString[0] - '0';
                    indexString = macroString.Substring(macroString.IndexOf(".") + 1, macroString.IndexOf("=") - (macroString.IndexOf(".") + 1));

                    string pattern = @"^" + indexNum.ToString() + @"." + Regex.Escape(indexString) + @"$";
                    checkString = macroString.Substring(0, macroString.IndexOf("=")); //0.이동

                    //dot pattern 때문에 고생함
                    if (checkRegex(checkString, pattern))
                    {
                        log.Debug("현재 반복문 넘버 : " + i + " , 작업 이름: " + macroString);
                        try
                        {

                            switch (indexNum)
                            {
                                case 0:
                                    //접속방식 : pc버전 / mobile버전

                                    break;
                                case 1:
                                    //url 접속 또는 이동
                                    string url = checkUrl(macroString.Substring(macroString.IndexOf("=") + 1));
                                    makeIeProcess(url);
                                    break;

                                case 2:
                                    //검색 기록 삭제
                                    clearHistory();
                                    break;

                                case 3:
                                    //검색
                                    string search = macroString.Substring(macroString.IndexOf("=") + 1);
                                    searchStart(search);

                                    break;

                                case 4:
                                    //게시글 or 버튼클릭
                                    //검색방식 선택
                                    //키워드검색=뉴스$블로그$부동산
                                    string keyword = macroString.Substring(macroString.IndexOf("=") + 1).Trim();
                                    string[] keywordNum = keyword.Split(new char[] { '$' });
                                    int keywordLength = keywordNum.Length;

                                    int randomIndex = getRandomIndex(keywordLength);
                                    string selectedKeyword = keywordNum[randomIndex];

                                    if (indexString.StartsWith("키워드"))
                                    {
                                        workerThread = new Thread(() =>
                                        {
                                            try
                                            {

                                                findByKeyword(selectedKeyword);
                                                Thread.Sleep(3000);

                                            }
                                            catch (Exception ex)
                                            {
                                                log.Debug(ex);
                                            }
                                        });
                                    }
                                    else
                                    {
                                        workerThread = new Thread(() =>
                                        {
                                            try
                                            {

                                                findByHref(selectedKeyword);
                                                Thread.Sleep(3000);
                                            }
                                            catch (Exception ex)
                                            {
                                                log.Debug(ex);
                                            }

                                        });
                                    }
                                    workerThread.SetApartmentState(ApartmentState.STA);
                                    workerThread.Start();
                                    workerThread.Join();

                                    changeIeTab();
                                    break;
                                case 5:
                                    //랜덤검색

                                    workerThread = new Thread(() =>
                                    {

                                        try
                                        {

                                            Console.WriteLine("clickRandomly start");

                                            clickRandomly();


                                            Thread.Sleep(3000);

                                        }
                                        catch (Exception ex)
                                        {
                                            log.Debug(ex);
                                        }

                                    });
                                    workerThread.SetApartmentState(ApartmentState.STA);
                                    workerThread.Start();
                                    workerThread.Join();


                                    changeIeTab();
                                    break;
                                case 6:
                                    //체류시간추가=30~50
                                    string st = macroString.Substring(macroString.IndexOf("=") + 1, macroString.IndexOf("~") - (macroString.IndexOf("=") + 1));
                                    int start_time = Int32.Parse(st);

                                    string et = macroString.Substring(macroString.IndexOf("~") + 1, (macroString.Length - 1) - macroString.IndexOf("~"));
                                    int end_time = Int32.Parse(et);

                                    sleep(start_time, end_time);
                                    break;
                                case 7:
                                    //마우스 이벤트 추가
                                    //마우스이벤트=왼쪽버튼 (더블)클릭:(100,300)
                                    string leftOrRight = macroString.Substring(macroString.IndexOf("=") + 1, macroString.IndexOf(" ") - (macroString.IndexOf("=") + 1));

                                    string clickOrDoubleClick = macroString.Substring(macroString.IndexOf(" ") + 1, macroString.IndexOf(":") - (macroString.IndexOf(" ") + 1));

                                    string xpos = macroString.Substring(macroString.IndexOf("(") + 1, macroString.IndexOf(",") - (macroString.IndexOf("(") + 1));
                                    string ypos = macroString.Substring(macroString.IndexOf(",") + 1, macroString.IndexOf(")") - (macroString.IndexOf(",") + 1));

                                    if (clickOrDoubleClick.StartsWith("더블"))
                                    {
                                        LeftDoubleClick(xpos, ypos);
                                    }
                                    else
                                    {
                                        LeftOneClick(xpos, ypos);
                                    }
                                    break;
                                default:
                                    Console.WriteLine("매크로형식이 맞지 않음");
                                    break;
                            }

                            //switch문 종료
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex);
                            log.Debug(ex);
                        }
                        finally
                        {
                            endCurrentMacro(macroString, i);
                        }
                    }
                    else
                    {
                        //regex가 맞지 않을때
                        //작업명령 format이 이상함.
                        continue;
                    }


                }
                splitInfiniteLoop(); // while문 반복마다 한번씩 실행됨 
                //deleteAllTab();
                ie = null;
                deleteAllIeProcesses();
                totalsw.Stop();

            }
        }

        void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 0;
            progressBar1.Style = ProgressBarStyle.Blocks;
            progressBar1.Value = progressBar1.Minimum;

            reportTextBox.Text += endAllMacro(ref totalsw);

            reportTextBox.SelectionStart = reportTextBox.TextLength;
            reportTextBox.ScrollToCaret();

            currentMacroLabel.Text = "현재 진행 명령 : 작업이 중단되었습니다.";

            processStartButton.Enabled = true;
            processStopButton.Enabled = true;
            if (e.Error != null)
            {
                reportTextBox.Text += "에러가 발생해서 작업이 중단되었습니다." + Environment.NewLine;
            }
            reportTextBox.Text += "작업을 중단했습니다." + Environment.NewLine;
            worker = null;
            deleteAllIeProcesses();

            ie = null;
            if (rk != null)
                rk.SetValue(null, "");
            if (aa != null)
            {
                aa.Abort();
                aa.Join();
                aa = null;
            }
            //예외처리 필요?


        }
        #endregion

        #region 마우스 좌표탐지
        //실시간 좌표 감지 시작
        private void mouseDetectStart(object sender, EventArgs e)
        {
            mouseDetectTimer = setTimer();
            mouseDetectTimer.Elapsed += timer_Elapsed;
            mouseDetectTimer.Start();
        }
        //실시간 좌표 감지 종료
        private void mouseDetectStop(object sender, EventArgs e)
        {
            mouseDetectTimer.Stop();
            mouseDetectTimer.Dispose();
            mouseDetectTimer = null;
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



        //ie 프로세스 생성
        private void makeIeProcess(string url)
        {
            /*
             * have to check resolution
             * basic resolution : 800 x 600 
             * */
            url = checkUrl(macroString.Substring(macroString.IndexOf("=") + 1));
            if (ie == null)
            {
                ie = new InternetExplorer();
                webBrowser = (SHDocVw.WebBrowser)ie;
                webBrowser.Visible = true;
                ie.Left = 0;
                ie.Top = 0;
                ie.Height = int.Parse(browserXSizeTextbox.Text);
                ie.Width = int.Parse(browserYSizeTextbox.Text);


            }
            //User-Agent: Mozilla / 5.0(Linux; U; Android 2.2) AppleWebKit / 533.1(KHTML, like Gecko) Version / 4.0 Mobile Safari/ 533.1"
            // "User-Agent: Mozilla/7.0(Linux; Android 7.0.0; SGH-i907) AppleWebKit/664.76 (KHTML, like Gecko) Chrome/87.0.3131.15 Mobile Safari/664.76 (Windows NT 10.0; WOW64; Trident/7.0; Touch; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729; Tablet PC 2.0; rv:11.0) like Gecko"

            log.Debug("navigate2 start");
            ie.Navigate2(url, null, null, null, null);

            log.Debug("navigate2 complete");
            ie.Wait();
            prevUrl = url;
            log.Debug("wait complete");

        }

        #region 탭 관리 기능
        //현재 탭을 제외한 모든 탭 제거
        private void deleteTab()
        {
            SHDocVw.ShellWindows shellWindows = new SHDocVw.ShellWindows();

            foreach (InternetExplorer iea in shellWindows)
            {
                if (iea.LocationURL.CompareTo(ie.LocationURL) != 0)
                {
                    iea.Quit();
                }
            }

        }

        //모든 탭 제거
        private void deleteAllTab()
        {
            SHDocVw.ShellWindows shellWindows = new SHDocVw.ShellWindows();
            foreach (InternetExplorer iea in shellWindows)
            {

                iea.Quit();


            }
        }

        //마지막 탭으로 ie 변경
        private void changeIeTab()
        {
            // Console.WriteLine("changeIeTab 시작");
            log.Debug("changeIeTab Start");
            ShellWindows allBrowsers = new ShellWindows();
            for (int i = allBrowsers.Count - 1; i >= 0; i--)
            {
                if (allBrowsers.Item(i) != null && !string.IsNullOrEmpty(((SHDocVw.InternetExplorer)allBrowsers.Item(i)).LocationURL))
                {
                    ie = (InternetExplorer)allBrowsers.Item(i);
                    webBrowser = (SHDocVw.WebBrowser)ie;
                    ie.Wait();
                    //   Console.WriteLine("changeIeTab 끝");
                    log.Debug("changeIeTab End");
                    return;
                }
            }
            //Uri uri = new Uri(url);

            //var shellWindows = new SHDocVw.ShellWindows();
            //int count = 0;
            //foreach (SHDocVw.InternetExplorer iex in shellWindows)
            //{
            //    //File Explorer인 경우 LocationURL가 비어있음. 제외.
            //    if (!string.IsNullOrEmpty(iex.LocationURL))
            //    {
            //Uri wbUri = new Uri(iex.LocationURL);
            //        Console.WriteLine(count +" , " + "wbUri : " + wbUri);
            //        if (wbUri.Equals(uri))
            //        {   
            //            Console.WriteLine("ie 변경완료 , ie 변경완료, ie 변경완료");
            //            ie = (InternetExplorer)shellWindows.Item((object)count);
            //            ie.Wait();
            //            return;
            //        }
            //    }
            //    count++;
            //}

            //Console.WriteLine("count : " + count);
            //SHDocVw.ShellWindows shellWindowsa = new SHDocVw.ShellWindows();

            //int length = shellWindowsa.Count;
            //object len = (object)(length - 1);

            //ie = (InternetExplorer)shellWindowsa.Item(len);
            //webBrowser = (SHDocVw.WebBrowser)ie;
            //ie.Wait();
            //Console.WriteLine(ie.LocationURL);
         //   Console.WriteLine("changeIeTab 반환 안되고 끝");
            log.Debug("changeIeTab 반환없이 End");
        }
        #endregion

        private void button1_Click_1(object sender, EventArgs e)
        {

            if (ie == null)
            {
                ie = new InternetExplorer();
                webBrowser = (SHDocVw.WebBrowser)ie;
                webBrowser.Visible = true;
            }

            ie.Navigate("https://m.sports.naver.com/news.nhn?oid=477&aid=0000232998");
            ie.Wait();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            searchMethodSelectCombobox.SelectedIndex = 0;

        }




        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (rk != null)
                rk.SetValue(null, "");
            if (aa != null)
            {
                aa.Abort();
                aa.Join();
                aa = null;
            }
        }
    }

    #region 로딩 완료 확장 메서드
    // 페이지 로딩 완료까지 대기하는 확장 메서드
    public static class SHDovVwEx
    {
        public static void Wait(this SHDocVw.InternetExplorer ie, int millisecond = 0)
        {
            int count = 0;
            while (ie.Busy == true || ie.ReadyState != SHDocVw.tagREADYSTATE.READYSTATE_COMPLETE)
            {
                System.Threading.Thread.Sleep(100);
                count++;
                if (count == 300) return;
            }

        }
    }
    #endregion

}