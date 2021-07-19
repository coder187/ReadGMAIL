using System;
using ReadEmails;


namespace ReadEmails_Caller
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Clear();
            Console.WriteLine("starting program...");

            //ReadEmail_Settings RS = new ReadEmail_Settings();
            ReadEmails.ReadEmail_Settings RS = new ReadEmails.ReadEmail_Settings();

            Console.ReadKey();
           
           
        }
    }
}
