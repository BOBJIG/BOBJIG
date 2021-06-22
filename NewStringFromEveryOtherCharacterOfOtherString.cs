using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication5
{
    class NewStringFromEveryOtherCharacterOfOtherString
    {
        public static void Main(string[] args)
        {
            string originalString, newString ="";
            Console.WriteLine("Enter the String");
            originalString = Console.ReadLine();
            for (int i = 0; i < originalString.Length; i= i + 2)
            {
                newString = newString + originalString[i];
            }
            Console.WriteLine("New string formed from odd characters of given string is: " + newString);
            Console.ReadLine();
        }
    }
}
