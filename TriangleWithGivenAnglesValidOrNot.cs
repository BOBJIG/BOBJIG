using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication5
{
    class TriangleWithGivenAnglesValidOrNot
    {
        public static void Main(string[] args)
        {
            int firstAngle, secondAngle, thirdAngle;
            Console.WriteLine("Enter first angle");
            firstAngle = Convert.ToInt32(Console.ReadLine());
            Console.WriteLine("Enter second angle");
            secondAngle = Convert.ToInt32(Console.ReadLine());
            Console.WriteLine("Enter third angle");
            thirdAngle = Convert.ToInt32(Console.ReadLine());
            if (firstAngle + secondAngle + thirdAngle == 180)
            {
                Console.WriteLine(" The triangle is valid");
            }
            else 
            Console.WriteLine(" The triangle is not valid");
            Console.ReadLine();
        }
    }
}
