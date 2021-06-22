using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication5
{
    class SumOfDigitsOfInteger
    {
        public static void Main(String[] args)
        {
            int intDigit, intNumber, intSum=0, intTemp;
            
            Console.WriteLine("Enter the Integer whose sum of digits is required");
            intNumber = Convert.ToInt32(Console.ReadLine());
            intTemp = intNumber;
            while (intNumber > 0)
            {
                intDigit = intNumber % 10;
                intSum = intSum + intDigit;
                intNumber = intNumber / 10;
            }
            Console.WriteLine("Sum of the Digits in Given Integer is: " +intSum);
            Console.ReadLine();
            }
        }
    }
