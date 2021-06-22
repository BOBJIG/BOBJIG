using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication5
{
    class CheckingBothNumbersAreEvenOrOdd
    {
        public static void Main(string[] args)
        {
            int firstNumber, secondNumber;
            bool bothNumbersAreEvenOrOdd = false;
            Console.WriteLine("Enter number 1");
            firstNumber = Convert.ToInt32(Console.ReadLine());
            Console.WriteLine("Enter number 2");
            secondNumber = Convert.ToInt32(Console.ReadLine());
            if ((firstNumber % 2 == 0 && secondNumber % 2 == 0) || (firstNumber % 2 == 1 && secondNumber % 2 == 1))
            {
                bothNumbersAreEvenOrOdd = true;
            }
            Console.WriteLine("Are both Numbers entered Even or Odd: " + bothNumbersAreEvenOrOdd);
            Console.ReadLine();
        }
    }
}
