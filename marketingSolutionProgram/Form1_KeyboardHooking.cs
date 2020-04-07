using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace marketingSolutionProgram
{
    partial class Form1
    {
        //[DllImport("User32.dll")]
        //static extern void keybd_event(byte vk, byte scan, int flags, int extra);

        //[DllImport("user32.dll")]
        //private static extern bool SetForegroundWindow(IntPtr hWnd);

        //const int KEYEVENTF_KEYDOWN = 0x00;
        //const int KEYEVENTF_KEYUP = 0x02;

        private void inputString(string inputString)
        {

            char[] idChars = inputString.ToCharArray();
            SetForegroundWindow((IntPtr)ie.HWND);

            foreach (char idChar in idChars)
            {
                if (idChar >= 'a' && idChar <= 'z')
                {
                    keybd_event((byte)(char.ToUpper(idChar)), 0, 0x00, 0);
                    keybd_event((byte)(char.ToUpper(idChar)), 0, 0x02, 0);

                }
                else if (idChar >= 'A' && idChar <= 'Z')
                {
                    keybd_event((int)Keys.LShiftKey, 0x00, 0x00, 0);

                    keybd_event((byte)(char.ToUpper(idChar)), 0, 0x00, 0);
                    keybd_event((byte)(char.ToUpper(idChar)), 0, 0x02, 0);

                    keybd_event((int)Keys.LShiftKey, 0x00, 0x02, 0);
                }
                else
                {
                    int nValue = 0;
                    bool bShift = false;
                    switch (idChar)
                    {
                        case '~': bShift = true; nValue = (int)Keys.Oemtilde; break;
                        case '_': bShift = true; nValue = (int)Keys.OemMinus; break;
                        case '+': bShift = true; nValue = (int)Keys.Oemplus; break;
                        case '{': bShift = true; nValue = (int)Keys.OemOpenBrackets; break;
                        case '}': bShift = true; nValue = (int)Keys.OemCloseBrackets; break;
                        case '|': bShift = true; nValue = (int)Keys.OemPipe; break;
                        case ':': bShift = true; nValue = (int)Keys.OemSemicolon; break;
                        case '"': bShift = true; nValue = (int)Keys.OemQuotes; break;
                        case '<': bShift = true; nValue = (int)Keys.Oemcomma; break;
                        case '>': bShift = true; nValue = (int)Keys.OemPeriod; break;
                        case '?': bShift = true; nValue = (int)Keys.OemQuestion; break;

                        case '!': bShift = true; nValue = (int)Keys.D1; break;
                        case '@': bShift = true; nValue = (int)Keys.D2; break;
                        case '#': bShift = true; nValue = (int)Keys.D3; break;
                        case '$': bShift = true; nValue = (int)Keys.D4; break;
                        case '%': bShift = true; nValue = (int)Keys.D5; break;
                        case '^': bShift = true; nValue = (int)Keys.D6; break;
                        case '&': bShift = true; nValue = (int)Keys.D7; break;
                        case '*': bShift = true; nValue = (int)Keys.D8; break;
                        case '(': bShift = true; nValue = (int)Keys.D9; break;
                        case ')': bShift = true; nValue = (int)Keys.D0; break;

                        case '`': bShift = false; nValue = (int)Keys.Oemtilde; break;
                        case '-': bShift = false; nValue = (int)Keys.OemMinus; break;
                        case '=': bShift = false; nValue = (int)Keys.Oemplus; break;
                        case '[': bShift = false; nValue = (int)Keys.OemOpenBrackets; break;
                        case ']': bShift = false; nValue = (int)Keys.OemCloseBrackets; break;
                        case '\\': bShift = false; nValue = (int)Keys.OemPipe; break;
                        case ';': bShift = false; nValue = (int)Keys.OemSemicolon; break;
                        case '\'': bShift = false; nValue = (int)Keys.OemQuotes; break;
                        case ',': bShift = false; nValue = (int)Keys.Oemcomma; break;
                        case '.': bShift = false; nValue = (int)Keys.OemPeriod; break;
                        case '/': bShift = false; nValue = (int)Keys.OemQuestion; break;

                        case '1': bShift = false; nValue = (int)Keys.D1; break;
                        case '2': bShift = false; nValue = (int)Keys.D2; break;
                        case '3': bShift = false; nValue = (int)Keys.D3; break;
                        case '4': bShift = false; nValue = (int)Keys.D4; break;
                        case '5': bShift = false; nValue = (int)Keys.D5; break;
                        case '6': bShift = false; nValue = (int)Keys.D6; break;
                        case '7': bShift = false; nValue = (int)Keys.D7; break;
                        case '8': bShift = false; nValue = (int)Keys.D8; break;
                        case '9': bShift = false; nValue = (int)Keys.D9; break;
                        case '0': bShift = false; nValue = (int)Keys.D0; break;

                        case ' ': bShift = false; nValue = (int)Keys.Space; break;
                        case '\x1b': bShift = false; nValue = (int)Keys.Escape; break;
                        case '\b': bShift = false; nValue = (int)Keys.Back; break;
                        case '\t': bShift = false; nValue = (int)Keys.Tab; break;
                        case '\a': bShift = false; nValue = (int)Keys.LineFeed; break;
                        case '\r': bShift = false; nValue = (int)Keys.Enter; break;

                        default:
                            bShift = false; nValue = 0; break;

                    }

                    if (nValue != 0)
                    {
                        // Caps Lock의 상태에 따른 대/소문자 처리
                        if (bShift)
                        {
                            keybd_event((int)Keys.LShiftKey, 0x00, KEYEVENTF_KEYDOWN, 0);
                            Thread.Sleep(30);
                        }

                        // Key 눌림 처리함.
                        //int nValue = Convert.ToInt32(chValue);
                        //int nValue = (int)Keys.Oemtilde;
                        keybd_event((byte)nValue, 0x00, KEYEVENTF_KEYDOWN, 0);
                        Thread.Sleep(30);
                        keybd_event((byte)nValue, 0x00, KEYEVENTF_KEYUP, 0);
                        Thread.Sleep(30);

                        // Caps Lock 상태를 회복함.
                        if (bShift)
                        {
                            keybd_event((int)Keys.LShiftKey, 0x00, KEYEVENTF_KEYUP, 0);
                            Thread.Sleep(30);
                        }
                    }
                }


            }
        }

    }
}
