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


        int indexNum = 0;
        int macroListTextboxCursor = 0;

        string indexString = String.Empty;
        string checkString = String.Empty;
        string macroString = String.Empty;
        string macroList = String.Empty;

        private Stopwatch totalsw = null;

        System.Timers.Timer mouseDetectTimer = null; //좌표 감지에 쓰이는 타이머
        Random random = null;
        Thread workerThread = null;

        InternetExplorer ie = null;
        SHDocVw.WebBrowser webBrowser = null;
        HtmlAgilityPack.HtmlDocument document;
        HtmlNodeCollection nodes;
        
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

        public Form1()
        {
            InitializeComponent();
            deleteAllIeProcesses();

            rk =
      Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Internet Settings\5.0\User Agent", true);
            rk.SetValue(null, "");
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
                Console.WriteLine(ie.LocationURL);
                document = new HtmlWeb().Load(ie.LocationURL);
                nodes = document.DocumentNode.SelectNodes("//iframe[@src]");

                if (nodes != null)
                    count = nodes.Count;

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
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

        public bool ExplorerMessage()
        {
            var hwnd = FindWindow("#32770", "Internet Explorer");
            if (hwnd != IntPtr.Zero)
            {
                SetForegroundWindow((IntPtr)ie.HWND);
               
                SendKeys.SendWait("{ESC}");
                return true;
            }
            return false;

        }
     
        //close alert when pop up alert
        public void ActivateAndClickOkButton()
        {
            // find dialog window with titlebar text of "Message from webpage"

            for (int i = 0; i < 6; i++)
            {

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

                    // find button on dialog window: classname = "Button", text = "OK"
                    var btn = FindWindowEx(hwnd, IntPtr.Zero, "Button", "취소");


                    if (btn != IntPtr.Zero)
                    {

                        // activate the button on dialog first or it may not acknowledge a click msg on first try
                        SendMessage(btn, WM_ACTIVATE, WA_ACTIVE, 0);
                        // send button a click message

                        SendMessage(btn, BM_CLICK, 0, 0);
                        return;
                    }
                    else
                    {
                        btn = FindWindowEx(hwnd, IntPtr.Zero, "Button", "확인");
                        if (btn != IntPtr.Zero)
                        {
                            // activate the button on dialog first or it may not acknowledge a click msg on first try
                            SendMessage(btn, WM_ACTIVATE, WA_ACTIVE, 0);
                            // send button a click message

                            SendMessage(btn, BM_CLICK, 0, 0);
                            return;
                        }
                        //btn = FindWindowEx(hwnd,IntPtr.Zero, "Button", )
                        //Interaction.MsgBox("button not found!");
                    }
                }
                else
                {
                    Thread.Sleep(1000);

                    //Interaction.MsgBox("window not found!");
                }
            }

        }

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

                            if (href != null && keywordElement.InnerText != null && href.CompareTo(keywordElement.InnerText.Replace("\r\n","")) == 0)
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
                Console.WriteLine(tagName+ "," + elements.length);
                //main document에서 조사
           
                foreach (IHTMLElement temp in elements)
                {
                    string tempString = temp.innerText;
                    if(tempString !=null )
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

                        Console.WriteLine("clcickㄷㅇㄷㅇ");
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


        public bool elementClick(int cnt ,int randomIndex, IHTMLElementCollection elements)
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


        public void clickRandomly()
        {

            if (isPresentIframe() == false)
            {
                //main document 에서 실행 <- iframe이 존재하지 않으므로
                mshtml.HTMLDocument doc = (mshtml.HTMLDocument)ie.Document;
                var elements = doc.getElementsByTagName("a");
                int randomIndex = getRandomIndex(elements.length);
                elementClick(0,randomIndex, elements);

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
                    elementClick(0,randomIndex, elements);
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

                    if (frame.document != null)
                        doc = (mshtml.HTMLDocument)frame.document;
                    else return;

                    var aTags = doc.getElementsByTagName("a");
                    randomIndex = getRandomIndex(aTags.length);
                    elementClick(0,randomIndex, aTags);
                }

            }

        }
       
        //캐시 삭제
        public void clearHistory()
        {
            deleteAllIeProcesses();
            ie = null;

            System.Diagnostics.Process.Start("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 4351").WaitForExit();
            Console.WriteLine("history clear d");
            log.Debug("캐시 삭제 완료");
        }

        #region ui변경 methods
        private void searchMethodSelectCombobox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (searchMethodSelectCombobox.SelectedIndex == 0)
            {
                inputKeywordLabel.Text = "키워드입력: ";
                inputKeywordLabel.Visible = true;
                inputKeywordTextBox.Visible = true;
            }
            else if (searchMethodSelectCombobox.SelectedIndex == 1)
            {
                inputKeywordLabel.Text = "href태그입력: ";
                inputKeywordLabel.Visible = true;
                inputKeywordTextBox.Visible = true;
            }
        }

        private void macroListClearButton_Click(object sender, EventArgs e)
        {
            macroListTextBox.Clear();
        }

        //vmware같은곳에서는 사용할 수 없음. ip address 얻어옴.
        private IPAddress LocalIPAddress()
        {
            if (!System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable())
            {
                return null;
            }
            IPHostEntry host = Dns.GetHostEntry(Dns.GetHostName());
            return host
                .AddressList
                .FirstOrDefault(ip => ip.AddressFamily == AddressFamily.InterNetwork);
        }


        //보고서 textbox에 쓸 함수 구현, runworkerCompleted에서 사용됨
        private StringBuilder endAllMacro(ref Stopwatch sw)
        {
            StringBuilder successString = new StringBuilder();
            /*
             * 현재 IP : ipv4 string
             * 작업한 시간: HH:MM:SS.mmmm
             * 작업 완료 시간: 2019/12/12 12:24:19 AM/PM, macroString
             * 작업완료.
             */
            successString.AppendLine("현재 IP: " + LocalIPAddress());
            successString.AppendLine("작업한 시간: " + sw.Elapsed);
            string cntTime = System.DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss tt, ");
            successString.AppendLine("작업 완료 시간: " + cntTime);
            successString.AppendLine("작업 완료");
            return successString;
        }

        //현재 매크로가 시작
        private void printCurrentMacro(string macroString, int num)
        {
            if (currentMacroLabel.InvokeRequired)
            {
                currentMacroLabel.BeginInvoke(new Action(() => { currentMacroLabel.Text = num + " , " + "현재 진행 명령 : " + macroString; }));
                return;
            }
        }

        //현재 매크로 완료
        private void endCurrentMacro(string macroString, int num)
        {
            if (reportTextBox.InvokeRequired)
            {
                reportTextBox.BeginInvoke(new Action(() =>
                {
                    reportTextBox.Text += num + " , " + "작업완료 : " + macroString.Substring(macroString.IndexOf(".") + 1) + "\r\n";

                    reportTextBox.SelectionStart = reportTextBox.TextLength;
                    reportTextBox.ScrollToCaret();

                    int pos = macroListTextBox.GetFirstCharIndexFromLine(macroListTextboxCursor);

                    if (pos > -1)
                    {
                        macroListTextBox.Select(pos, 0);
                        macroListTextBox.ScrollToCaret();
                    }
                    macroListTextboxCursor++;

                }));
                return;
            }
        }

        //무한 반복문에서 경계문으로 사용
        private void splitInfiniteLoop()
        {
            if (reportTextBox.InvokeRequired)
            {
                reportTextBox.BeginInvoke(new Action(() =>
                {
                    reportTextBox.Text += "-------------------------------------------"
+ Environment.NewLine;
                    reportTextBox.SelectionStart = reportTextBox.TextLength;
                    reportTextBox.ScrollToCaret();
                    macroListTextboxCursor = 0;
                }));
            }
        }

        //매크로 추가 버튼
        private void addMacroButton_Click(object sender, EventArgs e)
        {
            string macroString = String.Empty;
            int selectedIndex = tabControl1.SelectedIndex;
            switch (selectedIndex)
            {
                case 0:
                    //버전
                    break;
                case 1:
                    //url 이동
                    macroString = "▶" + selectedIndex + ".이동=" + inputUrlTextbox.Text;
                    break;
                case 2:
                    //검색기록삭제
                    macroString = "▶" + selectedIndex + ".검색기록삭제=";
                    break;
                case 3:
                    //검색어
                    if (searchTextBox.Text == "")
                    {
                        MessageBox.Show("검색어를 입력하세요.");
                    }
                    else
                        macroString = "▶" + selectedIndex + ".검색어입력=" + searchTextBox.Text;
                    break;
                case 4:
                    //게시글 or 버튼 클릭
                    if (searchMethodSelectCombobox.SelectedIndex == 0)
                    {
                        macroString = "▶" + selectedIndex + ".키워드검색=" + inputKeywordTextBox.Text;
                    }
                    else
                    {
                        macroString = "▶" + selectedIndex + ".href검색=" + inputKeywordTextBox.Text;
                    }
                    break;
                case 5:
                    //랜덤검색
                    macroString = "▶" + selectedIndex + ".랜덤태그검색=";
                    break;
                case 6:
                    //체류시간 설정
                    if (sleepStartTextBox.Text == "" || sleepEndTextBox.Text == "")
                    {
                        MessageBox.Show("체류 시간을 입력하세요.");
                    }
                    else
                    {
                        macroString = "▶" + selectedIndex + ".체류시간추가=" + sleepStartTextBox.Text + "~" + sleepEndTextBox.Text;
                    }
                    break;
                case 7:
                    //마우스 이벤트
                    if (mouseEventComboBox.SelectedItem == null)
                    {
                        MessageBox.Show("이벤트를 선택해주세요.");
                    }
                    else
                    {
                        //마우스이벤트=왼쪽/오른쪽 클릭:(100,300)
                        if (xAbsLocTextBox.Text.Length == 0 || yAbsLocTextBox.Text.Length == 0)
                        {
                            MessageBox.Show("X좌표와 Y좌표를 입력해주세요.");
                        }
                        else
                        {
                            macroString = "▶" + selectedIndex + ".마우스이벤트=" + mouseEventComboBox.Text + ":(" +
                                xAbsLocTextBox.Text + "," + yAbsLocTextBox.Text + ")";
                        }
                    }
                    break;


            }
            macroListTextBox.Text += macroString + "\r\n";
        }


        private void urlClearButton_Click(object sender, EventArgs e)
        {
            inputUrlTextbox.Clear();
        }
        #endregion

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
            rk =
          Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Internet Settings\5.0\User Agent", true);
            if (radioButton2.Checked == true)
            {
                rk.SetValue(null, "Mozilla/5.0(Linux; U; Android 7.0.0; SGH-i907) AppleWebKit / 533.1(KHTML, like Gecko) Version / 4.0 Mobile Safari/ 533.1 Chrome/87.0.3131.15");
                if (rk != null)
                {
                    log.Debug("레지스트리 수정 완료");
                    
                 };
            }
            else
                rk.SetValue(null, "");
            while (true)
            {
                totalsw = setTotalSw();
                totalsw.Start();
                //beforenavigate
                Thread aa = new Thread(() =>
                {
                    while (true)
                    {

                        ActivateAndClickOkButton();

                        Thread.Sleep(2000);
                    }
                });
                aa.IsBackground = true;
                aa.Start();
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
                                    ckIE();
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

                                    ckIE();
                                    break;
                                case 5:
                                    //랜덤검색

                                    workerThread = new Thread(() =>
                                    {

                                        try
                                        {

                                            Console.WriteLine("clickRandomly start");

                                            clickRandomly();
                                            // clickRandomlyExceptOnclick();
                                            Console.WriteLine("clickRandomly complete");
                                            Thread.Sleep(3000);
                                            Console.WriteLine("thread complete");
                                        }
                                        catch (Exception ex)
                                        {
                                            log.Debug(ex);
                                        }

                                    });
                                    workerThread.SetApartmentState(ApartmentState.STA);
                                    workerThread.Start();
                                    workerThread.Join();
                                    Console.WriteLine("workerThread end");
                                    //  ActivateAndClickOkButton();

                                    ckIE();
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
                deleteAllIeProcesses();
                ie = null;
  
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

            if (e.Error != null)
            {
                reportTextBox.Text += "에러가 발생해서 작업이 중단되었습니다." + Environment.NewLine;
            }
            reportTextBox.Text += "작업을 중단했습니다." + Environment.NewLine;
            worker = null;
            deleteAllIeProcesses();
            
            ie = null;
            if(rk!=null) 
            rk.SetValue(null, "");

            //예외처리 필요?


        }
        #endregion

        #region 마우스 이벤트
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

        #region 검색어 입력

        private void searchStart(string search)
        {
            //mobile버전 / pc버전
            string prefixUrl = ie.LocationURL.Substring(8);
            //수정해야함
            Console.WriteLine("prefixUrl : " + prefixUrl);
            string[] splitUrl = prefixUrl.Split(new char[] { '.' });
            if (splitUrl[0].StartsWith("m"))
            {
                //모바일 버전
                string website = splitUrl[1];
                if (website.StartsWith("naver"))
                {
                    Console.WriteLine("mobileNaverSearch");
                    mobileNaverSearchbarStart(search);
                }
                else
                {
                    mobileDaumSearchbarStart(search);
                }
            }
            else
            {
                //pc버전
                string website = splitUrl[1];
                if (website.StartsWith("naver"))
                {
                    pcNaverSearchbarStart(search);
                }
                else
                {
                    pcDaumSearchbarStart(search);
                }
            }
        }

        private void pcNaverSearchbarStart(string search)
        {
            //find search_bar

            mshtml.HTMLDocument doc = (mshtml.HTMLDocument)webBrowser.Document;
            //   mshtml.HTMLDocument doc = webBrowser.Document;
            var searchBar = doc.getElementById("query");
            searchBar.setAttribute("value", search);

            var searchButton = doc.getElementById("search_btn");
            searchButton.click();
            //searchButton.InvokeMember("submit");
            //searchButton.click();

        }

        public static void KeyDown(int keycode)
        {

            keybd_event((byte)Keys.HanguelMode, 0, 0, 0);

            keybd_event((byte)keycode, 0, (int)KeyFlag.KE_DOWN, 0);

        }
        public static void KeyDown2(int keycode)
        {
            keybd_event((byte)keycode, 0, (int)KeyFlag.KE_DOWN, 0);

        }

        private void mobileNaverSearchbarStart(string search)
        {
            //mshtml.HTMLDocument doc = ie.Document;

            mshtml.HTMLDocument doc = (mshtml.HTMLDocument)ie.Document;

            try
            {
                var fakeSearchBar = doc.getElementById("MM_SEARCH_FAKE") as mshtml.IHTMLElement2;
                IHTMLElementCollection forms = null;
                if (fakeSearchBar == null)
                {
                    fakeSearchBar = doc.getElementById("sch_w") as mshtml.IHTMLElement2;
                    forms = fakeSearchBar.getElementsByTagName("form");
                    Console.WriteLine(forms.length);

                }
                if (fakeSearchBar != null)
                    fakeSearchBar.focus();
                else
                    forms.item(0).focus();

                IntPtr iePtr = (IntPtr)ie.HWND;
                SetFocus(iePtr);
                //ShowWindowAsync(iePtr, SW_SHOWMAXIMIZED);

                // 윈도우에 포커스를 줘서 최상위로 만든다
                SetForegroundWindow(iePtr);

                // KeyDown((int)Keys.A);
                // KeyDown2((int)Keys.K);
                //// KeyDown((int)Keys.S);

                //StringBuilder st = new StringBuilder("여수누수");
                //SendMessage(iePtr, WM_SETTEXT, IntPtr.Zero, st);
                //Thread.Sleep(3000);
                //Thread.Sleep(3000);
                ////Process ieProcess = System.Diagnostics.Process.GetProcessById(ie.HWND);
                //SetForegroundWindow((IntPtr)ie.HWND);

                ////SendMessage(textBox1.Handle, WM_GETTEXTLENGTH, IntPtr.Zero, IntPtr.Zero);

                //// 260 바이트 만큼 메모리 공간을 할당한다.
                //IntPtr textPtr = Marshal.AllocHGlobal(260);

                //// 텍스트박스1 의 텍스트를 textPtr 에 저장한다.
                ////SendMessage(textBox1.Handle, WM_GETTEXT, new IntPtr(260), textPtr);

                //// 포인터를 유니코드 문자열로 변환한다. 
                //String text1 = Marshal.PtrToStringUni(textPtr, 260);
                //int ieHandle = ie.HWND;

                //// 텍스트박스2 에 텍스트박스1 의 텍스트를 입력한다
                //SendMessage((IntPtr)ie.HWND, WM_SETTEXT, IntPtr.Zero, "여수누수");

                //// 사용이 끝난 포인터는 메모리에서 해제해준다.
                //Marshal.FreeHGlobal(textPtr);

                IHTMLElement realSearchBar = doc.getElementById("query");
                realSearchBar.setAttribute("value", search);

                IHTMLElementCollection buttons = doc.getElementsByTagName("button");

                foreach (IHTMLElement button in buttons)
                {
                    string buttonClass = null;

                    buttonClass = button.className;
                    //Console.WriteLine(classSpan);
                    if (buttonClass != null && buttonClass.CompareTo("sch_submit MM_SEARCH_SUBMIT") == 0)
                    {
                        Console.WriteLine("class correct");
                        button.click();
                        return;
                    }
                }
                Console.WriteLine(webBrowser.LocationURL);
                //   mshtml.HTMLDocument doc = webBrowser.Document;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }

            //mshtml.IHTMLElement searchBar = webBrowser.Document.GetElementById("MM_SEARCH_FAKE");
            //searchBar.setAttribute("value", search);
            //searchBar.setAttribute 
            //searchBar.parentElement.click();
            //Console.WriteLine(searchBar.parentElement.parentElement.outerHTML);

        }
        private void pcDaumSearchbarStart(string search)
        {

            mshtml.HTMLDocument doc = (mshtml.HTMLDocument)ie.Document;
            var searchBar = doc.getElementById("q");
            searchBar.setAttribute("value", search);

            HtmlAgilityPack.HtmlDocument doc2 = new HtmlWeb().Load(webBrowser.LocationURL);
            var searchButton = doc2.DocumentNode.SelectSingleNode("//*[@id='daumSearch']/fieldset/div/div/button[2]");

            var buttons = doc.getElementsByTagName("button");
            foreach (IHTMLElement button in buttons)
            {
                string buttonClass = button.className;
                if (buttonClass != null && searchButton.Attributes["class"].Value.CompareTo(button.className) == 0)
                {
                    Console.WriteLine("class correct");
                    button.click();
                    return;
                }
            }


        }
        private void mobileDaumSearchbarStart(string search)
        {
            try
            {
                mshtml.HTMLDocument doc = (mshtml.HTMLDocument)ie.Document;
                var searchBar = doc.getElementById("q");
                searchBar.setAttribute("value", search);

                HtmlAgilityPack.HtmlDocument doc2 = new HtmlWeb().Load(webBrowser.LocationURL);
                var searchButton = doc2.DocumentNode.SelectSingleNode("//*[@id='form_totalsearch']/fieldset/div/div/button[3]");

                var buttons = doc.getElementsByTagName("button");
                foreach (IHTMLElement button in buttons)
                {
                    String buttonClass = button.className;
                    if (buttonClass != null && searchButton.Attributes["class"].Value.CompareTo(button.className) == 0)
                    {

                        button.click();
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        #endregion

        Microsoft.Win32.RegistryKey rk = null;
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
          //  "User-Agent: Mozilla/7.0(Linux; Android 7.0.0; SGH-i907) AppleWebKit/664.76 (KHTML, like Gecko) Chrome/87.0.3131.15 Mobile Safari/664.76 (Windows NT 10.0; WOW64; Trident/7.0; Touch; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729; Tablet PC 2.0; rv:11.0) like Gecko"

            log.Debug("navigate2 start");
            ie.Navigate2(url, null, null, null, null);

            log.Debug("navigate2 complete");
            ie.Wait();
            log.Debug("wait complete");

        }

        private void processStopButton_Click(object sender, EventArgs e)
        {
            currentMacroLabel.Text = "현재 진행 명령 : 작업을 중단 중입니다. 현재 명령까지 실행함.";
            worker.CancelAsync();

        }

        private void processStartButton_Click(object sender, EventArgs e)
        {
            macroList = String.Empty;
            macroList = macroListTextBox.Text;
            List<string> splitMacroList = new List<string>(macroList.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries));
            int macroListLength = splitMacroList.Count;

            //작업 명령이 하나도 입력되지 않았을 때
            if (macroListLength == 0)
            { MessageBox.Show("작업 명령이 작성되지 않았습니다.."); return; }

            //접속 명령어가 제일 먼저 시작하지 않으면 시작이 안되게 해야함.
            //if (!splitMacroList[0].StartsWith("▶0.이동"))
            //{ MessageBox.Show("첫 작업으로 이동명령을 추가해야합니다."); return; }
            //else
            //{
            log.Debug("\n\n\n");
            processStartButton.Enabled = false;
            progressBar1.Style = ProgressBarStyle.Marquee;
            progressBar1.MarqueeAnimationSpeed = 50;
            worker = new BackgroundWorker();
            worker.DoWork += (obj, ev) => bw_DoWork(splitMacroList, macroListLength, ev);
            worker.WorkerSupportsCancellation = true;
            worker.RunWorkerCompleted += bw_RunWorkerCompleted;
            worker.RunWorkerAsync();
            //}
        }

        private void ckIE()
        {
            Console.WriteLine("ckIE 시작");
            log.Debug("ckIE Start");
            SHDocVw.ShellWindows shellWindows = new SHDocVw.ShellWindows();
            int length = shellWindows.Count;
            object len = (object)(length - 1);

            ie = (InternetExplorer)shellWindows.Item(len);
            webBrowser = (SHDocVw.WebBrowser)ie;
            ie.Wait();
            Console.WriteLine("ckIE 끝");
            log.Debug("ckIE End");
        }

        //탭 삭제 기능
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
            findByKeyword("좋아요 평가하기");
            findByKeyword("좋아요 평가하기");


            // Thread workerThread = new Thread(() =>
            //{


            //    if (ie == null)
            //    {
            //        ie = new InternetExplorer();
            //        webBrowser = (SHDocVw.WebBrowser)ie;
            //        webBrowser.Visible = true;
            //    }

            //    ie.Left = 0;
            //    ie.Top = 0;
            //    ie.Height = 800;
            //    ie.Width = 1000;
            //    ie.BeforeNavigate2 += explorer_BeforeNavigate2;
            //    //webBrowser = (SHDocVw.WebBrowser)ie;
            //    //backgroundthread로 실행
            //    ie.Navigate2("https://bns.plaync.com/update/history/2019/191204_complete?utm_source=naver&utm_medium=timeboard&utm_campaign=pre_cre2&utm_content=bs_naver_pc_timeboard_pre_cre2_200215_20t#main1");
            //    ie.Wait();

            //    clickRandomly();
            //    clickRandomly();

            //    // ie.DocumentComplete += webComplete;
            //    try
            //    {

            //        //clearHistory();
            //        //printCurrentMacro("navigate2");

            //        Thread.Sleep(5000);
            //        //SendMessage(textBox1.Handle, WM_GETTEXTLENGTH, IntPtr.Zero, IntPtr.Zero);

            //        // 260 바이트 만큼 메모리 공간을 할당한다.
            //        IntPtr textPtr = Marshal.AllocHGlobal(260);

            //        // 텍스트박스1 의 텍스트를 textPtr 에 저장한다.
            //        //SendMessage(textBox1.Handle, WM_GETTEXT, new IntPtr(260), textPtr);

            //        // 포인터를 유니코드 문자열로 변환한다. 
            //        String text1 = Marshal.PtrToStringUni(textPtr, 260);

            //        //int ieHandle = ie.HWND;
            //        //// 텍스트박스2 에 텍스트박스1 의 텍스트를 입력한다.
            //        ////SendMessage((IntPtr)ieHandle, WM_SETTEXT, IntPtr.Zero, textPtr);

            //        //// 사용이 끝난 포인터는 메모리에서 해제해준다.
            //        //Marshal.FreeHGlobal(textPtr);

            //        //Console.WriteLine(ie.HWND);
            //        //Console.WriteLine(webBrowser.LocationURL);
            //        //ie.Navigate2("https://www.daum.net");
            //        //ie.Wait();
            //        //Console.WriteLine(webBrowser.LocationURL);
            //        //Console.WriteLine(webBrowser.Document);

            //        // ckIE();
            //        //printCurrentMacro("뉴스판");

            //        //   findByKeyword("연예판");
            //        ie.Wait();
            //        //ckIE();
            //        // findByKeyword("베스트");
            //        ie.Wait();


            //        deleteTab();
            //        //findByHref("https://news.naver.com/");
            //    }
            //    catch (Exception ex)
            //    {
            //        Console.WriteLine(ex);
            //    }

            //});
            // workerThread.SetApartmentState(ApartmentState.STA);
            // workerThread.Start();



        }
      
        private void Form1_Load(object sender, EventArgs e)
        {
            searchMethodSelectCombobox.SelectedIndex = 0;

        }

        //해상도 적용 버튼 이벤트
        private void browserSizeButton_Click(object sender, EventArgs e)
        {
            MessageBox.Show("해상도 적용 완료 : " + browserXSizeTextbox.Text + " X " + browserYSizeTextbox.Text);
        }


        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
          if(rk !=null)
            rk.SetValue(null, "");
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