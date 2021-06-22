using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication5
{
    class DisplayingLettersInReverseOrder
    {
        public static void Main(string[] args)
        {
            int sizeOfArray;
            int i;
            Console.WriteLine("Enter the size of array");
            sizeOfArray = Convert.ToInt32(Console.ReadLine());
            string[] array = new string[sizeOfArray];
            Console.WriteLine("Enter the elements of Array");
            for (i = 0; i < sizeOfArray; i++)
            {
                array[i] = Console.ReadLine();
            }
            Console.WriteLine("Reverse Order of the elements is: ");
            for (i = sizeOfArray-1; i >= 0; i--)
            {
                Console.WriteLine(array[i] + ' ');
            }
            Console.ReadLine();

        }
    }
}
