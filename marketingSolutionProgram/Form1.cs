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
        string macroString = String.Empty;
        Random random;
        InternetExplorer ie;
        SHDocVw.WebBrowser webBrowser;

        private Random setRandomInstance()
        {
            if (random == null)
                random = new Random(Guid.NewGuid().GetHashCode());
            return random;
        }

        private void sleep(int s, int e)
        {
            random = setRandomInstance();
            int randomSecond = random.Next(s, e);
            Thread.Sleep(randomSecond * 1000);
            return;
        }

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

        public Form1()
        {
            InitializeComponent();

        }

        //check in main document
        public bool isInMainDocument(string keyword, bool isHref)
        {
            bool result = false;
            
            if(isHref)
            {
                //href 검색
                
            } else
            {
                //keyword 검색
            }

            return result;
        }

        //check iframe length
        public int countIframe()
        {
            int count = 0;
            HtmlAgilityPack.HtmlDocument docu = null;
            HtmlNodeCollection nodes = null; ;
            try
            {

                docu = new HtmlWeb().Load("https://www.naver.com/");
                Thread.Sleep(3000); //로딩 될 떄까지 기다려야 함.
                nodes = docu.DocumentNode.SelectNodes("//iframe[@src]");
                count = nodes.Count;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
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
            var ress = doc.DocumentNode.SelectSingleNode("//*[text()[contains(., '" + keyword + "' )]]");
            string tagName = ress.Name;

            IHTMLElementCollection elements = ie.Document.getElementsByTagName(tagName);
            Console.WriteLine("findByKeyword length : " + elements.length);

            foreach (IHTMLElement temp in elements)
            {
                if (temp.innerText != null && temp.innerText.CompareTo(ress.InnerText) == 0)
                {
                    temp.click();
                    break;
                }
            }
        }

        public void checkHrefString(string keyword)
        {
            //href="https://post.naver.com/viewer/postView.nhn?volumeNo=27373987&memberNo=17369166"
            if (keyword.StartsWith("href"))
            {

            }
        }

        //체크 메소드를 따로 빼서 try해주자. getAttribute부분

        public void findByHref(string keyword)
        {
            Console.WriteLine("findByHref start");
            mshtml.HTMLDocument doc = ie.Document;
            var elements = doc.getElementsByTagName("a");
            //href = "https://news.naver.com/"

            foreach (IHTMLElement temp in elements)
            {
                string hrefString = string.Empty;
                try
                {

                    //hrefString = temp.getAttribute("href");
                    hrefString = temp.innerText;
                }
                catch
                {
                    Console.WriteLine("aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa");
                }
                if (hrefString != null)
                {
                    Console.WriteLine(hrefString);
                }
                if (hrefString != null && hrefString.CompareTo(keyword) == 0)
                {
                    Console.WriteLine("findByHref good");
                    //temp.click();
                }
            }

            // extracting all links
            //foreach (HtmlNode link in doc.DocumentNode.SelectNodes("//a[@href]"))
            //{
            //    HtmlAttribute att = link.Attributes["href"];

            //    if (att.Value.Contains("a"))
            //    {
            //        // showing output
            //        Console.WriteLine(att.Value);
            //    }
            //}

        }
        public void clickRandomly()
        {

        }
        private void button1_Click(object sender, EventArgs e)
        {
            ie = new InternetExplorer();

            webBrowser = (SHDocVw.WebBrowser)ie;

            string url = @"http://www.naver.com";
            webBrowser.Visible = true;
            webBrowser.Navigate(url);
            Thread.Sleep(5000);
            Console.WriteLine(webBrowser.LocationURL);

            HtmlAgilityPack.HtmlDocument docu = null;
            HtmlNodeCollection nodes = null; ;
            try
            {

                docu = new HtmlWeb().Load("https://www.naver.com/");
                Thread.Sleep(3000); //로딩 될 떄까지 기다려야 함.
                nodes = docu.DocumentNode.SelectNodes("//iframe[@src]");

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }


            foreach (var node in nodes)
            {
                try
                {
                    HtmlAttribute attr = node.Attributes["src"];
                    Console.WriteLine(attr.Value);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                }

            }
            
            Stopwatch sw = new Stopwatch();
            sw.Start();
            mshtml.HTMLDocument doc = ie.Document;
            var elements = doc.getElementsByTagName("img");
            sw.Stop();
            Console.WriteLine("img tag 갯수 == " + elements.length + "    " + sw.ElapsedMilliseconds);
            //foreach(IHTMLElement temp in elements)
            //{
            //    if(temp !=null )
            //    {
            //        Console.WriteLine(temp.outerHTML);
            //    }
            //}
            Random random = new Random(Guid.NewGuid().GetHashCode());
            int randomLength = random.Next(0, elements.length);
            int cnt = 0;
            Thread.Sleep(2000);
            //네이버페이
            //webBrowser2.Document.Body.InnerText
            string href = "https://news.naver.com/";

            findByKeyword("G마켓");
            // var doc2 = new HtmlAgilityPack.HtmlDocument();
            HtmlAgilityPack.HtmlDocument doc2 = new HtmlWeb().Load("http://www.naver.com/");
            //  doc2.Load(url);
            //doc2.DocumentNode.SelectNodes
            //var ress = doc2.DocumentNode.SelectSingleNode("//*[text()[contains(., '네이버페이')]]");
            // Console.WriteLine(ress.Name);
            //var val = ress.Attributes["href"].Value; //
            Thread.Sleep(2000);
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

            List<string> folderPathList = new List<string> { _tempInetFile, _cookies, _history };

            foreach (string dirpath in folderPathList)
                try
                {
                    Directory.Delete(dirpath, true);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
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
                    break;
                case 4:
                    //게시글 or 버튼 클릭
                    break;
                case 5:
                    //랜덤검색
                    macroString = "▶" + selectedIndex + ".랜덤태그검색=";

                    break;
                case 6:
                    //체류시간 설정
                    break;

            }
            macroListTextbox.Text += macroString + "\r\n";
        }
    }
}