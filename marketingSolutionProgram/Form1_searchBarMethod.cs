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
            Console.WriteLine("prefixUrl : " + prefixUrl);
            log.Debug("prefixUrl : " + prefixUrl + " , ie.LocationURL : " + ie.LocationURL);
            Console.WriteLine("ie.LocationUrl : " + ie.LocationURL);
            Console.WriteLine("webbrowser.LocationUrl : " + webBrowser.LocationURL);
            //수정해야함
            if (radioButton1.Checked)
            {
                //pc
                if (prefixUrl.StartsWith("www"))
                {
                    if (prefixUrl.Contains("naver"))
                    {
                        log.Debug("pcNaverSearchbarStart");
                        Console.WriteLine("pcNaverSearchbarStart");
                        pcNaverSearchbarStart(search);

                    }
                    else
                    {
                        log.Debug("pcDaumSearchbarStart");
                        Console.WriteLine("pcDaumSearchbarStart");
                        pcDaumSearchbarStart(search);

                    }
                }
                else
                {
                    if (prefixUrl.Contains("naver"))
                    {
                        log.Debug("pcNaverSearchbarStart2");
                        Console.WriteLine("pcNaverSearchbarStart2");

                        pcNaverSearchbarStart2(search);

                    }
                    else
                    {
                        log.Debug("pcDaumSearchbarStart2");
                        Console.WriteLine("pcDaumSearchbarStart2");

                        pcDaumSearchbarStart2(search);

                    }
                }
            }
            else
            {
                //mobile version
                if (prefixUrl.Contains("search"))
                {
                    if (prefixUrl.Contains("naver"))
                    {
                        log.Debug("SearchbarStart2");
                        Console.WriteLine("mobileNaverSearchbarStart2");

                        mobileNaverSearchbarStart2(search);
                    }
                    else
                    {
                        log.Debug("mobileDaumSearchbarStart2");
                        Console.WriteLine("mobileDaumSearchbarStart2");

                        mobileDaumSearchbarStart2(search);
                    }
                }
                else
                {
                    if (prefixUrl.Contains("naver"))
                    {
                        log.Debug("mobileNaverSearchbarStart");
                        Console.WriteLine("mobileNaverSearchbarStart");

                        mobileNaverSearchbarStart(search);
                    }
                    else
                    {
                        log.Debug("mobileDaumSearchbarStart");
                        Console.WriteLine("mobileDaumSearchbarStart");

                        mobileDaumSearchbarStart(search);
                    }
                }

            }

        }

        private void pcNaverSearchbarStart(string search)
        {
            //find search_bar
            SetForegroundWindow((IntPtr)ie.HWND);

            mshtml.HTMLDocument doc = (mshtml.HTMLDocument)ie.Document;
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
            SetForegroundWindow((IntPtr)ie.HWND);

            mshtml.HTMLDocument doc = (mshtml.HTMLDocument)ie.Document;
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
                    log.Debug("button clicked");
                    button.click();
                    return;
                }
            }

        }

        //main화면에서 검색하고나서
        private void mobileNaverSearchbarStart2(string search)
        {
                SetForegroundWindow((IntPtr)ie.HWND);

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
                        log.Debug("button clicked");
                        Console.WriteLine("class correct");
                        button.click();
                        return;
                    }
                }

        }

        //mobile main ui
        private void mobileNaverSearchbarStart(string search)
        {
            //mshtml.HTMLDocument doc = ie.Document;


                SetForegroundWindow((IntPtr)ie.HWND);
                mshtml.HTMLDocument doc = (mshtml.HTMLDocument)ie.Document;
                var fakeSearchBar = doc.getElementById("MM_SEARCH_FAKE") as mshtml.IHTMLElement2;
                IHTMLElementCollection forms = null;
                if (fakeSearchBar == null)
                {
                    var tempFakeSearchBar = doc.getElementById("sch_w") as mshtml.IHTMLElement2;
                    forms = tempFakeSearchBar.getElementsByTagName("form");
                    Console.WriteLine(forms.length);

                }
                if (fakeSearchBar != null)
                    fakeSearchBar.focus();
                else
                    forms.item(0).focus();

                Thread.Sleep(1000);

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
                        log.Debug("button clicked");
                        Console.WriteLine("class correct");
                        button.click();
                        return;
                    }
                }



                //IntPtr iePtr = (IntPtr)ie.HWND;
                //SetFocus(iePtr);
                //ShowWindowAsync(iePtr, SW_SHOWMAXIMIZED);

                // 윈도우에 포커스를 줘서 최상위로 만든다
                //        SetForegroundWindow(iePtr);


                //   mshtml.HTMLDocument doc = webBrowser.Document;
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
                log.Debug(ex);
            }
        }
        private void mobileDaumSearchbarStart2(string search)
        {
            try
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
            catch (Exception ex)
            {
                log.Debug(ex);
                Console.WriteLine(ex);
            }
        }
        #endregion
    }
}
