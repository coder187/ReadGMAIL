using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ReadEmail_DLL;

namespace Caller
{
    class Program
    {
        static void Main(string[] args)
        {
            ProgramSettings MySettings = new ProgramSettings();
            MySettings.ClearAccessToken = false;
            MySettings.CredentialPath = "token.json";
            MySettings.IncSpamTrash = false;
            MySettings.UnReadOnly = true;
            MySettings.NoOfEmailsToRead = 3;


            //.NET Caller
            //ReadEmail RE = new ReadEmail();
            //List<Enquiry> Es = RE.ReadEmails(MySettings);
            //int i = 1;
            //foreach (Enquiry e in Es)
            //{
            //    WriteEnquiry(e, i);
            //    i++;
            //}

            //.COM Interop VBA / VB6 Caller
            ReadEmail RE;
            RE = new ReadEmail();
            RE.ReadEmails(MySettings);
            int i = 1;
            foreach (Enquiry e in RE) {
                WriteEnquiry(e,i);
                i++;
            }
            
            //List<Enquiry> Es = RE.ReadEmails(MySettings);
            //int i = 1;
            //foreach (Enquiry e in Es)
            //{
            //    WriteEnquiry(e, i);
            //    i++;
            //}


            Console.Read();

            void WriteEnquiry(Enquiry e, int ii)
            {
                Console.ForegroundColor = ConsoleColor.Magenta;
                Console.Write(ii);
                Console.ForegroundColor = ConsoleColor.Gray;
                Console.WriteLine("////////////////////////////////////////////////");
                Console.WriteLine(e.Acc);
                Console.WriteLine(e.Source);
                Console.WriteLine(e.Name);
                Console.WriteLine(e.Phone);
                Console.WriteLine(e.Email);
                Console.WriteLine(e.Bus);
                Console.WriteLine(e.Pickup);
                Console.WriteLine(e.Dest);
                Console.WriteLine(e.PickDate.ToString("dd/MM/yyyy"));
                Console.WriteLine(e.Return);
                Console.WriteLine("//////////////////////////////////////////////////");
            }

            
        }
    }
   
}
