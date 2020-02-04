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

namespace marketingSolutionProgram
{
    public partial class Form1 : Form
    {
        BackgroundWorker worker;

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
        private ILog log = LogManager.GetLogger("Program");
        System.Timers.Timer mouseDetectTimer = null; //좌표 감지에 쓰이는 타이머
        Random random = null;

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

        private void sleep(int s, int e)
        {
            random = setRandomInstance();
            int randomSecond = random.Next(s, e);
            Thread.Sleep(randomSecond * 1000);
            return;
        }

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

        }

        //check iframe length
        public bool isPresentIframe()
        {
            int count = 0;

            try
            {
                document = new HtmlWeb().Load(webBrowser.LocationURL);
                nodes = document.DocumentNode.SelectNodes("//iframe[@src]");
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

        public bool compareInIframe()
        {
            bool result = false;
            return result;
        }

        public int getIframeIndex()
        {
            int count = 0;
            return count;
        }

        public void findByKeyword(string keyword)
        {
            // declaring & loading dom
            HtmlWeb web = new HtmlWeb();
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc = web.Load(webBrowser.LocationURL);
            var keywordElement = doc.DocumentNode.SelectSingleNode("//*[text()[contains(., '" + keyword + "' )]]");
            string tagName = null;

            if (keywordElement == null)
            {
                //mindocument에 없는 것임
                if (isPresentIframe() == true)
                {
                    //int iframeCount = getIframeCount();

                    mshtml.HTMLDocument doc1 = (mshtml.HTMLDocument)webBrowser.Document;
                    Console.WriteLine(doc1.frames.length);
                    for (int i = 0; i < doc1.frames.length; i++)
                    {
                        object index = i;
                        mshtml.IHTMLWindow2 frame = (mshtml.IHTMLWindow2)doc1.frames.item(ref index);
                        doc1 = (mshtml.HTMLDocument)frame.document;

                        var htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        htmlDoc.LoadHtml(doc1.documentElement.outerHTML);
                        var ress = htmlDoc.DocumentNode.SelectSingleNode("//*[text()[contains(., '" + keyword + "' )]]");
                        if (ress != null) tagName = ress.Name;

                        Console.WriteLine(tagName + "이거 맞아? ");
                        if (tagName == null) Console.WriteLine("tagName is null");
                        else { Console.WriteLine(ress.GetAttributeValue("href", "")); }

                        var elements = doc1.getElementsByTagName("a");

                        foreach (IHTMLElement element in elements)
                        {
                            string href = element.innerText;
                            if (href != null) Console.WriteLine("href : " + href);
                            if (href != null && href.CompareTo(ress.InnerText) == 0)
                            {

                                element.click();
                                webBrowser.Refresh();
                                Thread.Sleep(5000);
                                Console.WriteLine(webBrowser.LocationURL);
                                return ;
                            }
                        }
                    }

                    //for (int i = 0; i < iframeCount; i++)
                    //{

                    //    mshtml.HTMLDocument doc1 = (mshtml.HTMLDocument)webBrowser.Document;
                    //    object index = i;
                    //    mshtml.IHTMLWindow2 frame = (mshtml.IHTMLWindow2)doc1.frames.item(ref index);
                    //    doc1 = (mshtml.HTMLDocument)frame.document;
                    //    Console.WriteLine(doc1.documentElement.innerHTML);

                    //    IHTMLElementCollection iframeElements = doc1.getElementsByTagName(tagName);
                    //    foreach (IHTMLElement temp in iframeElements)
                    //    {
                    //        if (temp.innerText != null && temp.innerText.CompareTo(keywordElement.InnerText) == 0)
                    //        {
                    //            Console.WriteLine("click " + webBrowser.LocationURL);
                    //            temp.click();
                    //            Thread.Sleep(3000);
                    //            return;
                    //        }
                    //    }
                    //    //for문 종료
                    //}
                    //if 문 내부 종료
                }
            }
            else
            {
                //maindocument에 존재
                tagName = keywordElement.Name;
                IHTMLElementCollection elements = ie.Document.getElementsByTagName(tagName);
                //main document에서 조사
                foreach (IHTMLElement temp in elements)
                {
                    if (temp.innerText != null && temp.innerText.CompareTo(keywordElement.InnerText) == 0)
                    {
                        temp.click();

                        return;
                    }
                }


            }




        }

        public void findByHref(string keyword)
        {
            mshtml.HTMLDocument doc = ie.Document;
            var elements = doc.getElementsByTagName("a");
            keyword = makeToHrefString(ref keyword);
            string hrefString = string.Empty;

            foreach (IHTMLElement temp in elements)
            {
                hrefString = temp.getAttribute("href");

                if (hrefString != null && hrefString.CompareTo(keyword) == 0)
                {
                    Console.WriteLine("findByHref good");
                    temp.click();
                    break;
                }
            }

        }

        //check in main document
        public bool isInMainDocument(string[] keyword, bool isHref)
        {
            bool result = false;
            int keywordLength = keyword.Length;
            int randomIndex = getRandomIndex(keywordLength);
            string selectedKeyword = keyword[randomIndex];

            if (isHref)
            {
                //href 검색
                findByHref(selectedKeyword);

            }
            else
            {
                //keyword 검색
                findByKeyword(selectedKeyword);
            }

            return result;
        }

        public bool isInIframe(string keyword)
        {
            bool result = false;


            return result;
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

        public void elementClick(int randomIndex, IHTMLElementCollection elements)
        {
            int count = 0;
            foreach (IHTMLElement element in elements)
            {
                if (count == randomIndex)
                {
                    element.click();
                    break;
                }
                count++;
            }
        }

        public void clickRandomly()
        {
            if (isPresentIframe() == false)
            {
                //main document 에서 실행
                mshtml.HTMLDocument doc = ie.Document;
                var elements = doc.getElementsByTagName("a");
                int randomIndex = getRandomIndex(elements.length);
                elementClick(randomIndex, elements);

            }
            else
            {

            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            ie = new InternetExplorer();
            webBrowser = (SHDocVw.WebBrowser)ie;

            //string url = @"http://www.naver.com";
            webBrowser.Visible = true;
            //webBrowser.Navigate(url);

            // https://blog.naver.com/kims_pr/221780431616
            string url = @"https://www.naver.com";
            ie.Navigate(url);
            ie.Wait();

            mshtml.HTMLDocument doc1 = (mshtml.HTMLDocument)webBrowser.Document;
            Console.WriteLine(doc1.frames.length);
            object index = 3;
            mshtml.IHTMLWindow2 frame = (mshtml.IHTMLWindow2)doc1.frames.item(ref index);
            doc1 = (mshtml.HTMLDocument)frame.document;
            //   Console.WriteLine(doc1.documentElement.outerHTML);

            var htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.LoadHtml(doc1.documentElement.outerHTML);

            var ress = htmlDoc.DocumentNode.SelectSingleNode("//*[text()[contains(., 'G마켓' )]]");
            string tagName = null;

            if (ress != null) tagName = ress.Name;
            Console.WriteLine(tagName + "이거 맞아? ");
            if (tagName == null) Console.WriteLine("tagName is null");
            else { Console.WriteLine(ress.GetAttributeValue("href", "")); }

            var elements = doc1.getElementsByTagName("a");
            foreach (IHTMLElement element in elements)
            {
                string href = element.innerText;
                if (href != null) Console.WriteLine("href : " + href);
                if (href != null && href.CompareTo(ress.InnerText) == 0)
                {

                    element.click();
                    webBrowser.Refresh();
                    Thread.Sleep(5000);
                    Console.WriteLine(webBrowser.LocationURL);
                    break;
                }
            }

            //   HtmlWeb web = new HtmlWeb();
            //  HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            // doc = web.Load(doc1.documentElement.outerHTML); //error



            // 위 url의 iframe src를 출력한다.
            /*
            findByKeyword("G마켓");

            HtmlAgilityPack.HtmlDocument docu = null;
            HtmlNodeCollection nodes = null; ;
            try
            {

                docu = new HtmlWeb().Load("https://blog.naver.com/kims_pr/221780431616");

                //add microsoft mshtml object library reference 
                mshtml.HTMLDocument doc = (mshtml.HTMLDocument)webBrowser.Document;
                object index = 0;
                mshtml.IHTMLWindow2 frame = (mshtml.IHTMLWindow2)doc.frames.item(ref index);
                doc = (mshtml.HTMLDocument)frame.document;
                Console.WriteLine(doc.documentElement.innerHTML);

                var elements = doc.getElementsByTagName("a");
                int elementsLength = elements.length;
                random = setRandomInstance();
                int randomIndex = random.Next(0, elementsLength);

                int i = 0;
                foreach (IHTMLElement element in elements)
                {
                    string href = element.getAttribute("href");
                    if (href != null && i == randomIndex) { Console.WriteLine("click"); element.click(); break; }
                    i++;

                }

                Thread.Sleep(3000); //로딩 될 떄까지 기다려야 함.

                //  nodes = docu.DocumentNode.SelectNodes("//iframe[@src]");

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }

    */
            //foreach (var node in nodes)
            //{
            //    try
            //    {
            //        HtmlAttribute attr = node.Attributes["src"];
            //        Console.WriteLine(attr.Value);
            //    }
            //    catch (Exception ex)
            //    {
            //        Console.WriteLine(ex);
            //    }

            //}

            //foreach(IHTMLElement temp in elements)
            //{
            //    if(temp !=null )
            //    {
            //        Console.WriteLine(temp.outerHTML);
            //    }
            //}
            //webBrowser2.Document.Body.InnerText
            // var doc2 = new HtmlAgilityPack.HtmlDocument();
            HtmlAgilityPack.HtmlDocument doc2 = new HtmlWeb().Load("http://www.naver.com/");
            //  doc2.Load(url);
            //doc2.DocumentNode.SelectNodes
            //var ress = doc2.DocumentNode.SelectSingleNode("//*[text()[contains(., '네이버페이')]]");
            // Console.WriteLine(ress.Name);
            //var val = ress.Attributes["href"].Value; //
            // Console.WriteLine(ress.OuterHtml);
            //< span class="an_txt">네이버페이</span>
            //foreach (IHTMLElement temp in elements)
            //{
            //    if (temp.innerText != null && temp.innerText.CompareTo("네이버페이") == 0)
            //    {
            //        Console.WriteLine("That's correct : " + temp.outerHTML);
            //        temp.click();
            //    }
            //}

        }

        public void clearHistory()
        {
            string _tempInetFile = Environment.GetFolderPath(Environment.SpecialFolder.InternetCache);
            string _cookies = Environment.GetFolderPath(Environment.SpecialFolder.Cookies);
            string _history = Environment.GetFolderPath(Environment.SpecialFolder.History);
            Console.WriteLine(_tempInetFile);
            System.IO.DirectoryInfo di = new DirectoryInfo(_tempInetFile);
            foreach (FileInfo file in di.GetFiles())
            {
                file.Delete();
            }
            foreach (DirectoryInfo dir in di.GetDirectories())
            {
                dir.Delete(true); //delete subdirectories and files
            }
            //List<string> folderPathList = new List<string> {  _cookies, _history };

            //foreach (string dirpath in folderPathList)
            //    try
            //    {
            //        Console.WriteLine(dirpath);
            //        Directory.Delete(dirpath, true);
            //    }
            //    catch (Exception ex)
            //    {
            //        Console.WriteLine(ex);
            //    }
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
        private void printCurrentMacro(string macroString)
        {
            if (currentMacroLabel.InvokeRequired)
            {
                currentMacroLabel.BeginInvoke(new Action(() => { currentMacroLabel.Text = "현재 진행 명령 : " + macroString; }));
                return;
            }
        }

        //현재 매크로 완료
        private void endCurrentMacro(string macroString)
        {
            if (reportTextBox.InvokeRequired)
            {
                reportTextBox.BeginInvoke(new Action(() =>
                {
                    reportTextBox.Text += "작업완료 : " + macroString.Substring(macroString.IndexOf(".") + 1);

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
                }));
            }
        }


        #endregion

        private string checkUrl(string macroString)
        {
            bool includeHttp = macroString.StartsWith("http");
            if (includeHttp == false)
            {
                return "https://" + macroString;
            }
            else
            {
                return macroString;
            }
        }

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

        bool checkRegex(string checkString, string pattern)
        {
            return Regex.IsMatch(checkString, pattern);
        }

        #region BackgroundWorker event define
        void bw_DoWork(List<string> splitMacroList, int macroListLength, DoWorkEventArgs e)
        {
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
                    printCurrentMacro(macroString); //현재 명령을 라벨에 출력

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
                                    if (ie == null)
                                    {
                                        ie = new InternetExplorer();
                                        webBrowser = (SHDocVw.WebBrowser)ie;
                                        webBrowser.Visible = true;
                                    }
                                    ie.Navigate(url);
                                    ie.Wait();
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
                                    if (indexString.StartsWith("키워드"))
                                    {
                                        //keyword_search(ref driver, keywordNum);
                                    }
                                    else
                                    {
                                        //searchLink(ref driver, keywordNum);
                                    }
                                    break;
                                case 5:
                                    //랜덤검색
                                    clickRandomly();
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
                        }
                        finally
                        {
                            endCurrentMacro(macroString);
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
            //if (e.Cancelled)
            //{
            //    reportTextbox.Text += "작업을 중단했습니다." + Environment.NewLine;
            //    worker = null;
            //    if(driver != null) {
            //        driver.Quit();
            //    driver = null;
            //    }
            //}
            if (e.Error != null)
            {
                reportTextBox.Text += "에러가 발생해서 작업이 중단되었습니다." + Environment.NewLine;
            }
            reportTextBox.Text += "작업을 중단했습니다." + Environment.NewLine;
            worker = null;
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
            string prefixUrl = string.Empty;
            //수정해야함
            Console.WriteLine("prefixUrl : " + prefixUrl);
            string[] splitUrl = prefixUrl.Split(new char[] { '.' });
            if (splitUrl[0].StartsWith("m"))
            {
                //모바일 버전
                string website = splitUrl[1];
                if (website.StartsWith("naver"))
                {
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

            mshtml.HTMLDocument doc = webBrowser.Document;
            //   mshtml.HTMLDocument doc = webBrowser.Document;
            var searchBar = doc.getElementById("query");
            searchBar.setAttribute("value", search);

            var searchButton = doc.getElementById("search_btn");
            searchButton.click();
            //searchButton.InvokeMember("submit");
            //searchButton.click();

        }
        private void mobileNaverSearchbarStart(string search)
        {
            //mshtml.HTMLDocument doc = ie.Document;

            mshtml.HTMLDocument doc = ie.Document;
            try
            {
                var fakeSearchBar = doc.getElementById("MM_SEARCH_FAKE") as mshtml.IHTMLElement2;

                fakeSearchBar.focus();

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

            mshtml.HTMLDocument doc = ie.Document;
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
                mshtml.HTMLDocument doc = ie.Document;
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
                        Console.WriteLine("class correct");
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

        private void urlClearButton_Click(object sender, EventArgs e)
        {
            inputUrlTextbox.Clear();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            clearHistory();
        }
    }


    #region 로딩 완료 확장 메서드
    // 페이지 로딩 완료까지 대기하는 확장 메서드
    public static class SHDovVwEx
    {
        public static void Wait(this SHDocVw.InternetExplorer ie, int millisecond = 0)
        {
            while (ie.Busy == true || ie.ReadyState != SHDocVw.tagREADYSTATE.READYSTATE_COMPLETE)
            {
                System.Threading.Thread.Sleep(100);
            }
            System.Threading.Thread.Sleep(millisecond);
        }
    }
    #endregion

}