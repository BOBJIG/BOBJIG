using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication5
{
    class MultiplicationOfThreeNumbers
    {
        static void Main(string[] args)
        {
            int firstNumber,secondNumber,thirdNumber,intResult;
            Console.WriteLine("Enter number 1");
            firstNumber= Convert.ToInt32( Console.ReadLine());
            Console.WriteLine("Enter number 2");
            secondNumber= Convert.ToInt32(Console.ReadLine());
            Console.WriteLine("Enter number 3");
            thirdNumber= Convert.ToInt32(Console.ReadLine());
            intResult = firstNumber* secondNumber* thirdNumber;
            Console.WriteLine("Multiplication of the numbers is: " + intResult);
            Console.ReadLine();
        }
    }
}
