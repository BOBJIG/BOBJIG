using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication5
{
    class TypeOfTriangleWithGivenAngles
    {
        public static void Main(string[] args)
        {
            int firstAngle, secondAngle, thirdAngle;
            string typeOfTriangle = "";
            Console.WriteLine("Enter first angle");
            firstAngle = Convert.ToInt32(Console.ReadLine());
            Console.WriteLine("Enter second angle");
            secondAngle = Convert.ToInt32(Console.ReadLine());
            Console.WriteLine("Enter third angle");
            thirdAngle = Convert.ToInt32(Console.ReadLine());
            if (firstAngle == secondAngle && secondAngle == thirdAngle)
                typeOfTriangle = "Equilateral";
            else if ((firstAngle == secondAngle ||firstAngle == thirdAngle ||secondAngle ==thirdAngle) && (firstAngle != thirdAngle||secondAngle!= thirdAngle||firstAngle !=secondAngle))
                typeOfTriangle = "Isosceles";
            else
                typeOfTriangle = "Scalene";
            Console.WriteLine("Type of traingle that can be formed with the given angles is: " + typeOfTriangle);
            Console.ReadLine();


        }
    }
}
