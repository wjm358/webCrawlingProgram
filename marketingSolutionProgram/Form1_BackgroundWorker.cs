using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace marketingSolutionProgram
{
    partial class Form1

     {
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
                        Console.WriteLine(ex);
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
                                    ie.Wait();
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
                                                Console.WriteLine(ex);
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
                                                Console.WriteLine(ex);

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
                                            Console.WriteLine(ex);

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
                reportTextBox.Text += e.Error.ToString() + Environment.NewLine;
                log.Debug("에러 발생 : " + e.Error.ToString());
                
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

    }
}
