using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication5
{
    class SumOfElementsInIntegerArray
    {
        public static void Main(string[] args)
        {
            int sizeOfArray, i,sumOfElements =0 ;
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
                sumOfElements = sumOfElements + firstArray[i];
            }
            Console.WriteLine("sum of the elements of Array is: "+ sumOfElements);
            Console.ReadLine();

            }
    }
}
