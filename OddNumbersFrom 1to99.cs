using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication5
{
    class OddNumbersFrom1to99
    {
        public static void Main(String[] args)
        {
            Console.WriteLine("Odd numbers from 1 to 99 are as follows:");
            for (int i = 1; i <= 99; i++)
            {
                if (i % 2 != 0)
                {
                    Console.WriteLine(i + "\n");
                }
            }
            Console.ReadLine();
        }
    }
}
