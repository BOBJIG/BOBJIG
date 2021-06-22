using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication5
{
    class OperationOnTwoNumbersAndDisplayingResult
    {
        
        public static void Main(string[] args)
        {
            int firstNumber, secondNumber;
            char operationToPerform;
            double Result=0;
            Console.WriteLine("Enter number 1");
            firstNumber = Convert.ToInt32(Console.ReadLine());
            Console.WriteLine("Enter number 2");
            secondNumber = Convert.ToInt32(Console.ReadLine());
            Console.WriteLine("Enter Operation to be Performed");
            operationToPerform = Convert.ToChar(  Console.ReadLine());
            switch (operationToPerform)
            {
                case '+':
                    Result = firstNumber + secondNumber;
                    break;
                case '-':
                    Result = firstNumber - secondNumber;
                    break;
                case '*':
                    Result = firstNumber * secondNumber;
                    break;
                case '/':
                    if (secondNumber == 0)
                        {
                            Console.WriteLine("Please enter the second number other than 0");
                            secondNumber = Convert.ToInt32(Console.ReadLine());
                        }
                    Result = firstNumber / secondNumber;
                    break;
            }
            Console.WriteLine(firstNumber + " " + operationToPerform + " " + secondNumber + " = " + Result);
            Console.ReadLine();
        }
    }
}
