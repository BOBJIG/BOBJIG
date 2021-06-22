using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication5
{
    class NoOfDaysInMonth
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("Enter month between 1-12 to get no of days");
            int month = Convert.ToInt32(Console.ReadLine());
            int noOfDays=0;
            if (!(month < 1 || month > 12))
            {

                if (month == 2)
                    noOfDays = 28;
                else if (month == 1 || month == 3 || month == 5 || month == 7 || month == 8 || month == 10 || month == 12)
                    noOfDays = 31;
                else 
                    noOfDays = 30;
                Console.WriteLine("Number of Days in given month is: " + noOfDays);
            }
            else
                 Console.WriteLine("Please enter valid Month");
             
                
             Console.ReadLine();
        }

        
	
	
    }
}
