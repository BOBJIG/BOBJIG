using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication5
{
    class MultiplyElementsOfArray
    {
        public static void Main(String[] args)
        {
            int sizeOfArray,i;
            Console.WriteLine("Enter the size of the array");
            sizeOfArray = Convert.ToInt32(Console.ReadLine());
            int[] firstArray = new int[sizeOfArray];
            int[] secondArray = new int[sizeOfArray];
            Console.WriteLine("Enter the elements of firstArray");
            for (i = 0; i < firstArray.Length; i++)
            {
                firstArray[i] = Convert.ToInt32(Console.ReadLine());
            }
            Console.WriteLine("Enter the elements of firstArray");
            for (i = 0; i < secondArray.Length; i++)
            {
                secondArray[i] = Convert.ToInt32(Console.ReadLine());
            }
            Console.WriteLine("Multiplication Result of two arrays is:");
            for (i = 0; i < sizeOfArray; i++)
            {
                Console.WriteLine(firstArray[i] * secondArray[i] +" ");
            }
            Console.ReadLine();
        }

    }
}
