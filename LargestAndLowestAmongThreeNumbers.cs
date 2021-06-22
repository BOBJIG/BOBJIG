using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication5
{
    class LargestAndLowestAmongThreeNumbers
    {
        public static void Main(String[] args)
        {
            int firstNumber, secondNumber, thirdNumber, largestNumber, lowestNumber;
            Console.WriteLine("Enter number 1");
            firstNumber = Convert.ToInt32(Console.ReadLine());
            Console.WriteLine("Enter number 2");
            secondNumber = Convert.ToInt32(Console.ReadLine());
            Console.WriteLine("Enter number 3");
            thirdNumber = Convert.ToInt32(Console.ReadLine());
            // To find Largest Number
            if (firstNumber > secondNumber && firstNumber > thirdNumber)
            {
                largestNumber = firstNumber;
            }
            else   if (secondNumber > thirdNumber)
                {
                    largestNumber = secondNumber;
                }
            else
                largestNumber = thirdNumber;
            Console.WriteLine("Largest Number is: " + largestNumber);

            //To find Lowest Number
            if (firstNumber < secondNumber && firstNumber < thirdNumber)
            {
                lowestNumber = firstNumber;
            }
            else if (secondNumber < thirdNumber)
            {
                lowestNumber = secondNumber;
            }
            else
                lowestNumber = thirdNumber;
                Console.WriteLine("Lowest Number is: " + lowestNumber);
            Console.ReadLine();

        }

    }
}
