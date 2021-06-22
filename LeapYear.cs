using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication5
{
    class LeapYear
    {
        public static void Main(string[] args)
        {
            bool isLeapYear = false;
            int yearToCheck;
            Console.WriteLine("Enter year to check whether it is leap year or not");
            yearToCheck = Convert.ToInt32(Console.ReadLine());
            if (((yearToCheck % 4 == 0) && (yearToCheck % 100 != 0)) || (yearToCheck % 400 == 0))
            {
                isLeapYear = true;
                
            }
            Console.WriteLine("Is Leap Year: " + isLeapYear);
            Console.ReadLine();
        }
        
    }
}
