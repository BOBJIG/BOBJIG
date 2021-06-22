using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication5
{
    class swapTwoNumbers
    {
        
        public static void Main(String[] args)
        {
            int firstNumber, secondNumber;
            Console.WriteLine("Enter number 1");
            firstNumber = Convert.ToInt32(Console.ReadLine());
            Console.WriteLine("Enter number 2");
            secondNumber = Convert.ToInt32(Console.ReadLine());
            Console.WriteLine("Before Swaping: firstNumber = " + firstNumber + " and secondNumber = " + secondNumber);
            firstNumber = firstNumber - secondNumber;
            secondNumber = firstNumber + secondNumber;
            firstNumber = secondNumber - firstNumber;
            Console.WriteLine("After Swaping: firstNumber = " + firstNumber + " and secondNumber = " + secondNumber);
            Console.ReadLine();

        }
         
    }
}
