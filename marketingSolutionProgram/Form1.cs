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
        HtmlAgilityPack.HtmlDocument document;
        HtmlNodeCollection nodes;

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
            var ress = doc.DocumentNode.SelectSingleNode("//*[text()[contains(., '" + keyword + "' )]]");
            string tagName = ress.Name;
            if(tagName == null)
            {
                //mindocument에 없는 것임
                var ress2 = doc.DocumentNode.SelectNodes("//iframe[@src]");
                Console.WriteLine(ress2[0].OuterHtml);
            }
            else
            {
                //maindocument에 존재
            }
            IHTMLElementCollection elements = ie.Document.getElementsByTagName(tagName);

            //main document에서 조사
            foreach (IHTMLElement temp in elements)
            {
                if (temp.innerText != null && temp.innerText.CompareTo(ress.InnerText) == 0)
                {
                    temp.click();

                    return;
                }
            }

            if (isPresentIframe() == true)
            {
                int iframeCount = getIframeCount();

                for (int i = 0; i < iframeCount; i++)
                {

                    mshtml.HTMLDocument doc1 = (mshtml.HTMLDocument)webBrowser.Document;
                    object index = i;
                    mshtml.IHTMLWindow2 frame = (mshtml.IHTMLWindow2)doc1.frames.item(ref index);
                    doc1 = (mshtml.HTMLDocument)frame.document;
                    Console.WriteLine(doc1.documentElement.innerHTML);

                    IHTMLElementCollection iframeElements = doc1.getElementsByTagName(tagName);
                    foreach (IHTMLElement temp in iframeElements)
                    {
                        if (temp.innerText != null && temp.innerText.CompareTo(ress.InnerText) == 0)
                        {
                            temp.click();

                            return;
                        }
                    }
                    //for문 종료
                }
                //if 문 내부 종료
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

        //check in main document
        public bool isInMainDocument(string keyword, bool isHref)
        {
            bool result = false;

            if (isHref)
            {
                //href 검색
                findByHref(keyword);

            }
            else
            {
                //keyword 검색
                findByKeyword(keyword);
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

            //href="https://post.naver.com/viewer/postView.nhn?volumeNo=27373987&memberNo=17369166"
            if (keyword.StartsWith("href"))
            {
                keyword = keyword.Substring(keyword.IndexOf("="));
            }
            return keyword;

        }

        //체크 메소드를 따로 빼서 try해주자. getAttribute부분


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

            string url = @"http://www.naver.com";
            webBrowser.Visible = true;
            //webBrowser.Navigate(url);

            // https://blog.naver.com/kims_pr/221780431616
            ie.Navigate(url);
            ie.Wait();

            // 위 url의 iframe src를 출력한다.

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
}