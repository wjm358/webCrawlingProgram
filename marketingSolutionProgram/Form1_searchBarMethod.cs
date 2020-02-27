using HtmlAgilityPack;
using mshtml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace marketingSolutionProgram
{
    partial class Form1
    {
        #region 검색어 입력

        private void searchStart(string search)
        {
            //mobile버전 / pc버전
            string prefixUrl = ie.LocationURL.Substring(8);
            //수정해야함
            if(radioButton1.Checked)
            {
                //pc
                if(prefixUrl.StartsWith("www"))
                {
                    if(prefixUrl.Contains("naver"))
                    {
                        pcNaverSearchbarStart(search);
                        ie.Wait();
                    }
                    else
                    {
                        pcDaumSearchbarStart(search);
                        ie.Wait();
                    }
                } else
                {
                    if(prefixUrl.Contains("naver"))
                    {
                        pcNaverSearchbarStart2(search);
                        ie.Wait();
                    }
                    else
                    {
                        pcDaumSearchbarStart2(search);
                        ie.Wait();
                    }
                }
            }
            else
            {
              if(prefixUrl.Contains("search")) {
                    if(prefixUrl.Contains("naver"))
                    {
                        mobileNaverSearchbarStart2(search);
                    }else
                    {
                        mobileDaumSearchbarStart2(search);
                    }
                }
              else
                {
                    if(prefixUrl.Contains("naver"))
                    {
                        mobileNaverSearchbarStart(search);
                    }
                    else
                    {
                        mobileDaumSearchbarStart(search);
                    }
                }

            }
            //string[] splitUrl = prefixUrl.Split(new char[] { '.' });
            //if (splitUrl[0].StartsWith("m"))
            //{
            //    //모바일 버전
            //    //string website = splitUrl[1];
            //    if (prefixUrl.Contains("naver"))
            //    {
            //        Console.WriteLine("mobileNaverSearch");
            //        if (prefixUrl.Contains("search")) mobileNaverSearchbarStart2(search);
            //        mobileNaverSearchbarStart(search);
            //    }
            //    else
            //    {
            //        if (prefixUrl.Contains("search")) mobileDaumSearchbarStart2(search);
            //        mobileDaumSearchbarStart(search);
            //    }
            //}
            //else
            //{
            //    //pc버전
            //    //string website = splitUrl[1];
            //    if (prefixUrl.Contains("naver"))
            //    {
            //        if (prefixUrl.Contains("search")) pcNaverSearchbarStart2(search);
            //        pcNaverSearchbarStart(search);
            //    }
            //    else
            //    {
            //        if (prefixUrl.Contains("search")) pcDaumSearchbarStart2(search);
            //        pcDaumSearchbarStart(search);
            //    }
            //}
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
        private void pcNaverSearchbarStart2(string search)
        {
           
            mshtml.HTMLDocument doc = (mshtml.HTMLDocument)webBrowser.Document;
            var searchBar = doc.getElementById("nx_query");
            searchBar.setAttribute("value", "");
            searchBar.setAttribute("value", search);

            IHTMLElementCollection buttons = doc.getElementsByTagName("button");

            foreach (IHTMLElement button in buttons)
            {
                string buttonClass = null;

                buttonClass = button.className;
                //Console.WriteLine(classSpan);
                if (buttonClass != null && buttonClass.CompareTo("bt_search") == 0)
                {
                    Console.WriteLine("class correct");
                    button.click();
                    return;
                }
            }

        }
        //public static void KeyDown(int keycode)
        //{

        //    keybd_event((byte)Keys.HanguelMode, 0, 0, 0);

        //    keybd_event((byte)keycode, 0, (int)KeyFlag.KE_DOWN, 0);

        //}
        //public static void KeyDown2(int keycode)
        //{
        //    keybd_event((byte)keycode, 0, (int)KeyFlag.KE_DOWN, 0);

        //}

         private void mobileNaverSearchbarStart2(string search)
        {
          
            mshtml.HTMLDocument doc = (mshtml.HTMLDocument)ie.Document;
            var fakeSearchBar = doc.getElementById("nx_query") as mshtml.IHTMLElement2;
            fakeSearchBar.focus();
            ((IHTMLElement)fakeSearchBar).setAttribute("value", "");
            Thread.Sleep(2000);
            ((IHTMLElement)fakeSearchBar).setAttribute("value", search);

            IHTMLElementCollection buttons = doc.getElementsByTagName("button");

            foreach (IHTMLElement button in buttons)
            {
                string buttonClass = null;

                buttonClass = button.className;
                //Console.WriteLine(classSpan);
                if (buttonClass != null && buttonClass.CompareTo("btn_search") == 0)
                {
                    Console.WriteLine("class correct");
                    button.click();
                    return;
                }
            }
           
        }
        private void mobileNaverSearchbarStart(string search)
        {
            //mshtml.HTMLDocument doc = ie.Document;

            mshtml.HTMLDocument doc = (mshtml.HTMLDocument)ie.Document;

            try
            {
                var fakeSearchBar = doc.getElementById("MM_SEARCH_FAKE") as mshtml.IHTMLElement2;
                //IHTMLElementCollection forms = null;
                //if (fakeSearchBar == null)
                //{
                //    fakeSearchBar = doc.getElementById("sch_w") as mshtml.IHTMLElement2;
                //    forms = fakeSearchBar.getElementsByTagName("form");
                //    Console.WriteLine(forms.length);

                //}
                if (fakeSearchBar != null)
                {
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

                } 
                //else
                //    forms.item(0).focus();

                //IntPtr iePtr = (IntPtr)ie.HWND;
                //SetFocus(iePtr);
                //ShowWindowAsync(iePtr, SW_SHOWMAXIMIZED);

                // 윈도우에 포커스를 줘서 최상위로 만든다
                //        SetForegroundWindow(iePtr);


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
        private void pcDaumSearchbarStart2(string search)
        {
           
            mshtml.HTMLDocument doc = (mshtml.HTMLDocument)ie.Document;
            var searchBar = doc.getElementById("q");
            ((IHTMLElement2)searchBar).clearAttributes();
            searchBar.setAttribute("value", "");
            searchBar.setAttribute("value", search);

            var button = doc.getElementById("daumBtnSearch");
            button.click();

        }

        private void mobileDaumSearchbarStart(string search)
        {
            try
            {
                mshtml.HTMLDocument doc = (mshtml.HTMLDocument)ie.Document;
                var searchBar = doc.getElementById("q");
                searchBar.setAttribute("value", search);

                IHTMLElementCollection buttons = doc.getElementsByTagName("button");

                foreach (IHTMLElement button in buttons)
                {
                    string buttonClass = null;

                    buttonClass = button.className;
                    //Console.WriteLine(classSpan);
                    if (buttonClass != null && buttonClass.CompareTo("btn_sch btn_search") == 0)
                    {
                        Console.WriteLine("class correct");
                        button.click();
                        return;
                    }
                }

                //HtmlAgilityPack.HtmlDocument doc2 = new HtmlWeb().Load(webBrowser.LocationURL);
                //var searchButton = doc2.DocumentNode.SelectSingleNode("//*[@id='form_totalsearch']/fieldset/div/div/button[3]");

                //var buttons = doc.getElementsByTagName("button");
                //foreach (IHTMLElement button in buttons)
                //{
                //    String buttonClass = button.className;
                //    if (buttonClass != null && searchButton.Attributes["class"].Value.CompareTo(button.className) == 0)
                //    {

                //        button.click();
                //        return;
                //    }
                //}
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
        private void mobileDaumSearchbarStart2(string search)
        {
          
            mshtml.HTMLDocument doc = (mshtml.HTMLDocument)ie.Document;
            var searchBar = doc.getElementById("q");
            ((IHTMLElement2)searchBar).clearAttributes();
            searchBar.setAttribute("value", "");
            searchBar.setAttribute("value", search);


            //btn_sch btn_search

            IHTMLElementCollection buttons = doc.getElementsByTagName("button");

            foreach (IHTMLElement button in buttons)
            {
                string buttonClass = null;

                buttonClass = button.className;
                //Console.WriteLine(classSpan);
                if (buttonClass != null && buttonClass.CompareTo("btn_sch btn_search") == 0)
                {
                    Console.WriteLine("class correct");
                    button.click();
                    return;
                }
            }

        }
        #endregion
    }
}
