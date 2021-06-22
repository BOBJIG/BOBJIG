using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication5
{
    class ElectricityBillOfGivenCustomer
    {
        public static void Main(string[] args)
        {
            int customerId, noOfUnits;
            string name;
            decimal chargePerUnit=0, surcharge=0, totalAmount=0;
            Console.WriteLine("Enter customer Id");
            customerId = Convert.ToInt32(Console.ReadLine());
            Console.WriteLine("Enter Name");
            name = Console.ReadLine();
            Console.WriteLine("Enter no. of Units");
            noOfUnits = Convert.ToInt32(Console.ReadLine());
            if (noOfUnits < 199)
                chargePerUnit = 1.20M;
            else if (noOfUnits >= 200 && noOfUnits < 400)
                chargePerUnit = 1.50M;
            else if (noOfUnits >= 400 && noOfUnits < 600)
                chargePerUnit = 1.80M;
            else if (noOfUnits >= 600)
                chargePerUnit = 2.00M;
            totalAmount = noOfUnits * chargePerUnit;
            if (totalAmount < 100)
                totalAmount = 100;
            surcharge = totalAmount * 15 / 100;
            Console.WriteLine("Customer Id: " + customerId);
            Console.WriteLine("Customer Name: " + name);
            Console.WriteLine("Units Consumed: " + noOfUnits);
            Console.WriteLine("Amount Charges @ Rs." + chargePerUnit + " per unit: " + totalAmount);
            Console.WriteLine("Surcharge Amount: " + surcharge);
            Console.WriteLine("Net amount paid by Customer: " + ((totalAmount+surcharge)));
            Console.ReadLine();
        }
    }
}
