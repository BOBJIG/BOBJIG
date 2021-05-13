using System;
using System.Collections.Generic;
using System.Text;

namespace FirstTestProject
{
    class Addition
    {
        public void Sum()
        {

             // 12345
            // 1+2+3+4+5= 15

             Console.WriteLine("enter the number digits");
             int digitsCount=Convert.ToInt16(Console.ReadLine());

             Console.WriteLine("enter the actual digits");
             int actualNumber = Convert.ToInt16(Console.ReadLine());


             int sum = 0;
            
            //while (i < digitsCount)
            //{
            //    int k = actualNumber % 10;
            //   int k1 = actualNumber / 10;
            //   actualNumber = k1;
            //   sum = sum + k;
            //   i++;

            //}
            for (int i = 0; i < digitsCount; i++)
            {
                int k = actualNumber % 10;
                int k1 = actualNumber / 10;
                actualNumber = k1;
                sum = sum + k;

            }
            Console.WriteLine($"the sum of the digits given is :{sum}");

            Console.ReadLine();


        }
    }
}
