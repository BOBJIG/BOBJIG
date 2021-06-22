using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication5
{
    class VowelOrConsonant
    {
        public static void Main(string[] args)
        {
            char checkVowelOrConsonant;
            bool isVowel = false;
            Console.WriteLine("Enter alphabet to check whether it is vowel or consonant");
            checkVowelOrConsonant = Convert.ToChar(Console.ReadLine());
            char[] array = { 'a', 'e', 'i', 'o', 'u' };
            for (int i = 0; i < array.Length; i++)
            {
                if (checkVowelOrConsonant == array[i])
                {
                    isVowel = true;
                    Console.WriteLine("The given alpahbet is Vowel"); 
                    break;
                }
            }
            if (isVowel == false)
            Console.WriteLine("The given alpahbet is Consonant");
            Console.ReadLine();

        }
    }
}
