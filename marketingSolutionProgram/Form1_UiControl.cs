using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace marketingSolutionProgram
{
    partial class Form1
    {
        //해상도 적용 버튼 이벤트
        private void browserSizeButton_Click(object sender, EventArgs e)
        {
            MessageBox.Show("해상도 적용 완료 : " + browserXSizeTextbox.Text + " X " + browserYSizeTextbox.Text);
        }
        private void processStopButton_Click(object sender, EventArgs e)
        {
            currentMacroLabel.Text = "현재 진행 명령 : 작업을 중단 중입니다. 현재 명령까지 실행함.";

            processStopButton.Enabled = false;

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
            {
                MessageBox.Show("작업 명령이 작성되지 않았습니다..");
                return;
            }

            log.Debug("\n\n\n");

            processStartButton.Enabled = false;
            progressBar1.Style = ProgressBarStyle.Marquee;
            progressBar1.MarqueeAnimationSpeed = 50;

            worker = new BackgroundWorker();
            worker.DoWork += (obj, ev) => bw_DoWork(splitMacroList, macroListLength, ev);
            worker.WorkerSupportsCancellation = true;
            worker.RunWorkerCompleted += bw_RunWorkerCompleted;
            worker.RunWorkerAsync();
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

    }
}
