using HtmlAgilityPack;
using log4net;
using mshtml;
using SHDocVw;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

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


        #endregion

        public int indexNum = 0;
        public int macroListTextboxCursor = 0;

        public  string indexString = String.Empty;
        public string checkString = String.Empty;
        public string macroString = String.Empty;
        public string macroList = String.Empty;
        public string prevUrl = String.Empty;

        public Stopwatch totalsw = null;
        public Microsoft.Win32.RegistryKey rk = null;

        public Random random = null;
        public Thread workerThread = null;
        public Thread aa = null;

        public InternetExplorer ie = null;
        public SHDocVw.WebBrowser webBrowser = null;
        System.Windows.Forms.WebBrowser webb;
        public HtmlAgilityPack.HtmlDocument document;
        public HtmlNodeCollection nodes;
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

        const int KEYEVENTF_KEYDOWN = 0x00;
        const int KEYEVENTF_KEYUP = 0x02;

        //Form1_KeyboardHooking으로 이동
        //private void inputString(string inputString)
        //{
        //    char[] idChars = inputString.ToCharArray();
        //    SetForegroundWindow((IntPtr)ie.HWND);

        //    foreach (char idChar in idChars)
        //    {
        //        if (idChar >= 'a' && idChar <= 'z')
        //        {
        //            keybd_event((byte)(char.ToUpper(idChar)), 0, 0x00, 0);
        //            keybd_event((byte)(char.ToUpper(idChar)), 0, 0x02, 0);

        //        }
        //        else if (idChar >= 'A' && idChar <= 'Z')
        //        {
        //            keybd_event((int)Keys.LShiftKey, 0x00, 0x00, 0);

        //            keybd_event((byte)(char.ToUpper(idChar)), 0, 0x00, 0);
        //            keybd_event((byte)(char.ToUpper(idChar)), 0, 0x02, 0);

        //            keybd_event((int)Keys.LShiftKey, 0x00, 0x02, 0);
        //        }
        //        else
        //        {
        //            int nValue = 0;
        //            bool bShift = false;
        //            switch (idChar)
        //            {
        //                case '~': bShift = true; nValue = (int)Keys.Oemtilde; break;
        //                case '_': bShift = true; nValue = (int)Keys.OemMinus; break;
        //                case '+': bShift = true; nValue = (int)Keys.Oemplus; break;
        //                case '{': bShift = true; nValue = (int)Keys.OemOpenBrackets; break;
        //                case '}': bShift = true; nValue = (int)Keys.OemCloseBrackets; break;
        //                case '|': bShift = true; nValue = (int)Keys.OemPipe; break;
        //                case ':': bShift = true; nValue = (int)Keys.OemSemicolon; break;
        //                case '"': bShift = true; nValue = (int)Keys.OemQuotes; break;
        //                case '<': bShift = true; nValue = (int)Keys.Oemcomma; break;
        //                case '>': bShift = true; nValue = (int)Keys.OemPeriod; break;
        //                case '?': bShift = true; nValue = (int)Keys.OemQuestion; break;

        //                case '!': bShift = true; nValue = (int)Keys.D1; break;
        //                case '@': bShift = true; nValue = (int)Keys.D2; break;
        //                case '#': bShift = true; nValue = (int)Keys.D3; break;
        //                case '$': bShift = true; nValue = (int)Keys.D4; break;
        //                case '%': bShift = true; nValue = (int)Keys.D5; break;
        //                case '^': bShift = true; nValue = (int)Keys.D6; break;
        //                case '&': bShift = true; nValue = (int)Keys.D7; break;
        //                case '*': bShift = true; nValue = (int)Keys.D8; break;
        //                case '(': bShift = true; nValue = (int)Keys.D9; break;
        //                case ')': bShift = true; nValue = (int)Keys.D0; break;

        //                case '`': bShift = false; nValue = (int)Keys.Oemtilde; break;
        //                case '-': bShift = false; nValue = (int)Keys.OemMinus; break;
        //                case '=': bShift = false; nValue = (int)Keys.Oemplus; break;
        //                case '[': bShift = false; nValue = (int)Keys.OemOpenBrackets; break;
        //                case ']': bShift = false; nValue = (int)Keys.OemCloseBrackets; break;
        //                case '\\': bShift = false; nValue = (int)Keys.OemPipe; break;
        //                case ';': bShift = false; nValue = (int)Keys.OemSemicolon; break;
        //                case '\'': bShift = false; nValue = (int)Keys.OemQuotes; break;
        //                case ',': bShift = false; nValue = (int)Keys.Oemcomma; break;
        //                case '.': bShift = false; nValue = (int)Keys.OemPeriod; break;
        //                case '/': bShift = false; nValue = (int)Keys.OemQuestion; break;

        //                case '1': bShift = false; nValue = (int)Keys.D1; break;
        //                case '2': bShift = false; nValue = (int)Keys.D2; break;
        //                case '3': bShift = false; nValue = (int)Keys.D3; break;
        //                case '4': bShift = false; nValue = (int)Keys.D4; break;
        //                case '5': bShift = false; nValue = (int)Keys.D5; break;
        //                case '6': bShift = false; nValue = (int)Keys.D6; break;
        //                case '7': bShift = false; nValue = (int)Keys.D7; break;
        //                case '8': bShift = false; nValue = (int)Keys.D8; break;
        //                case '9': bShift = false; nValue = (int)Keys.D9; break;
        //                case '0': bShift = false; nValue = (int)Keys.D0; break;

        //                case ' ': bShift = false; nValue = (int)Keys.Space; break;
        //                case '\x1b': bShift = false; nValue = (int)Keys.Escape; break;
        //                case '\b': bShift = false; nValue = (int)Keys.Back; break;
        //                case '\t': bShift = false; nValue = (int)Keys.Tab; break;
        //                case '\a': bShift = false; nValue = (int)Keys.LineFeed; break;
        //                case '\r': bShift = false; nValue = (int)Keys.Enter; break;

        //                default:
        //                    bShift = false; nValue = 0; break;

        //            }

        //            if (nValue != 0)
        //            {
        //                // Caps Lock의 상태에 따른 대/소문자 처리
        //                if (bShift)
        //                {
        //                    keybd_event((int)Keys.LShiftKey, 0x00, KEYEVENTF_KEYDOWN, 0);
        //                    Thread.Sleep(30);
        //                }

        //                // Key 눌림 처리함.
        //                //int nValue = Convert.ToInt32(chValue);
        //                //int nValue = (int)Keys.Oemtilde;
        //                keybd_event((byte)nValue, 0x00, KEYEVENTF_KEYDOWN, 0);
        //                Thread.Sleep(30);
        //                keybd_event((byte)nValue, 0x00, KEYEVENTF_KEYUP, 0);
        //                Thread.Sleep(30);

        //                // Caps Lock 상태를 회복함.
        //                if (bShift)
        //                {
        //                    keybd_event((int)Keys.LShiftKey, 0x00, KEYEVENTF_KEYUP, 0);
        //                    Thread.Sleep(30);
        //                }
        //            }
        //        }


        //    }
        //}


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

        /* 
         *  todo list : 
         * */
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

                            //SetForegroundWindow(alertHwnd);
                            var buttonHwnd = FindWindowEx(alertHwnd, IntPtr.Zero, "Button", null);
                            if (buttonHwnd != IntPtr.Zero)
                            {
                                //Console.WriteLine("not zero");
                            SendMessage(buttonHwnd, WM_ACTIVATE, WA_ACTIVE, 0);
                            SendMessage(buttonHwnd, BM_CLICK, 0, 0);
                               return;
                            }

                        }
                    }
                }
            }

        }

        #endregion

        public Form1()
        {
            SetForegroundWindow((IntPtr)ie.HWND);
            InitializeComponent();
            rk =
         Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Internet Settings\5.0\User Agent", true);
            rk.SetValue(null, "");
            //uri 문제가 아님.
            //deleteAllIeProcesses();
      
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
            try
            {

                Uri uri = new Uri(ie.LocationURL);
                
            Console.WriteLine("ie.LocationURL : " + ie.LocationURL);
            document = new HtmlWeb().Load(uri.ToString()); //problem
                
            nodes = document.DocumentNode.SelectNodes("//iframe[@src]");

            if (nodes != null)
                count = nodes.Count;

            } catch(Exception ex)
            {
                log.Debug("Error in isPresentIfram method : " + ex + " , ie.LocationURL : " + ie.LocationURL);
                Console.WriteLine(ex + " , ie.LocationURL : " + ie.LocationURL);
            }



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

                } else
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
                Console.WriteLine("메인 document에서 수행gg");
            }
            else
            {
                //iframe이 존재
                //http://rpp.gmarket.co.kr/?exhib=33275 <<- about:blank 제외해야함
                int randomIndex = getRandomIndex(2);
                bool mainDocumentIsChoiced = (randomIndex == 0 ? false : true);

                if (mainDocumentIsChoiced == true)
                {
                    Console.WriteLine("메인 document에서 수행");
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
                    Console.WriteLine("iframe에서 수행");
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
                // webb = (IWebBrowserApp)ie;

                ie.Left = 0;
                ie.Top = 0;
                ie.Height = int.Parse(browserXSizeTextbox.Text);
                ie.Width = int.Parse(browserYSizeTextbox.Text);
                //ie.DocumentComplete += new DWebBrowserEvents2_DocumentCompleteEventHandler(printUrl);
               // ie.DocumentComplete += IE_DocumentComplete;
                //ie.NavigateComplete2 += IE_NavigateComplete2;
            }
            //User-Agent: Mozilla / 5.0(Linux; U; Android 2.2) AppleWebKit / 533.1(KHTML, like Gecko) Version / 4.0 Mobile Safari/ 533.1"
            // "User-Agent: Mozilla/7.0(Linux; Android 7.0.0; SGH-i907) AppleWebKit/664.76 (KHTML, like Gecko) Chrome/87.0.3131.15 Mobile Safari/664.76 (Windows NT 10.0; WOW64; Trident/7.0; Touch; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729; Tablet PC 2.0; rv:11.0) like Gecko"

            log.Debug("navigate2 start");
            ie.Navigate(url, null, null, null, null);

            log.Debug("navigate2 complete");
            ie.Wait();
           
            log.Debug("wait complete");

        }
        //void IE_DocumentComplete(object pDisp, ref object URL)
        //{
        //    Console.WriteLine("DocumentComplete : " + URL.ToString());
        //}

        //void IE_NavigateComplete2(object pDisp, ref object URL)
        //{
        //    Console.WriteLine("NavigateComplate : " + URL);
        //}

        public void printUrl(object pDisp, ref object URL)
        {
            Console.WriteLine("printUrl :" +(String)URL);
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
            log.Debug("changeIeTab 반환없이 End");
        }
        #endregion

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