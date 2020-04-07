using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace marketingSolutionProgram
{
    partial class Form1
    {
         public void inputId(string idString)
        {
            Console.WriteLine(idString.Length);

            char[] idChar = idString.ToCharArray();

            for (int i=0; i<idChar.Length;i++)
            {
                char value = idChar[i];
                Console.WriteLine("value : " + value);

            }

        } 
        public void inputPassword(string passwordString)
        {

        }
        public void checkChar(char a)
        {

        }
        public void pressKey(char a)
        {

        }
        public void randomSleep()
        {

            random = setRandomInstance();
            int randomSleepValue = random.Next(100, 300);
            Thread.Sleep(randomSleepValue);
            return;

        }
    }
}
