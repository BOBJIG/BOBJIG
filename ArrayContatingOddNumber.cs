using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication5
{
    class ArrayContatingOddNumber
    {
        public static void Main(string[] args)
        {
            int sizeOfArray, i;
            bool containsOddNumber = false;
            Console.WriteLine("Enter the size of the array");
            sizeOfArray = Convert.ToInt32(Console.ReadLine());
            int[] firstArray = new int[sizeOfArray];
            Console.WriteLine("Enter the elements of Array");
            for (i = 0; i < firstArray.Length; i++)
            {
                firstArray[i] = Convert.ToInt32(Console.ReadLine());
            }
            for (i = 0; i < firstArray.Length; i++)
            {
                if (firstArray[i] % 2 != 0)
                {
                    containsOddNumber = true;
                    break;
                }
            }
            Console.WriteLine("Array Contains Odd Number: " + containsOddNumber);
            Console.ReadLine();
        }
    }
}
