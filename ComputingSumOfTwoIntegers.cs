using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication5
{
    class ComputingSumOfTwoIntegers
    {
        
        public static void Main(String[] args)
        {
            int firstNumber, secondNumber;
            Console.WriteLine("Enter number 1");
            firstNumber = Convert.ToInt32(Console.ReadLine());
            Console.WriteLine("Enter number 2");
            secondNumber = Convert.ToInt32(Console.ReadLine());
            if (firstNumber == secondNumber)
            {
                Console.WriteLine("Both numbers are equal. Triple of their sum is: " + 3 *(firstNumber + secondNumber));
            }
            else
            {
                Console.WriteLine("Sum of two numbers is: " + (firstNumber + secondNumber));
            }
            Console.ReadLine();
        }

    }
}
