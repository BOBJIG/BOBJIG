using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication5
{
    class TriangleOfNumberWithSpecificWidth
    {

        public static void Main(string[] args)
        {
            int numberToDisplay, width, i;
            Console.WriteLine("Enter a number to be displayed");
            numberToDisplay = Convert.ToInt32(Console.ReadLine());
            Console.WriteLine("Enter Width");
            width = Convert.ToInt32(Console.ReadLine());
            i=width;
            if (width != 0)
            {
                for (i = width; i >0; i--)
                {
                    
                    for (int j = i; j > 0; j--)
                    {
                        Console.Write(numberToDisplay);
                    }
                    Console.WriteLine("");
                }
               
            }
            Console.ReadLine();

        }
    }
}
