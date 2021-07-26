//14072021
//jkelly@jsd.ie
//Jonathan Kelly
//coder187@hotmail.com
//program to read email or emails from given gmail inbox or email id and then 
//perist to db or otherwise make data available to caller..

using System;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Collections;

namespace ReadEmail_DLL
{
    public class Enquiry
    {
        public string Name { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }
        public string Bus { get; set; }
        public string Pickup { get; set; }
        public DateTime PickDate { get; set; }
        public string Dest { get; set; }
        public string Return { get; set; }
        public string Source { get; set; }
        public string EnquiryDateFormat { get; set; }
        public string Acc { get; set; }
    }

    public class ProgramSettings
    {
        public string CredentialPath { get; set; }
        public bool ClearAccessToken { get; set; }
        public int NoOfEmailsToRead { get; set; }
        public bool IncSpamTrash { get; set; }
        public bool UnReadOnly { get; set; }
        public string Filter_From { get; set; }
        public string Filter_Text { get; set; }
    }
    public class Pragma
    {
        public string BigString { get; set; }
        public string SmallString { get; set; }

    }
    //interface to allow vba read IEnumerable
    //https://limbioliong.wordpress.com/2011/10/28/exposing-an-enumerator-from-managed-code-to-com/

    public interface TestIEnumerableClass : IEnumerable
    {
        new IEnumerator GetEnumerator();
        void BuildList(Pragma p);
    }

    //test class for VBA interop
    [ClassInterface(ClassInterfaceType.None)]
    public class TestClass : TestIEnumerableClass
    {
       
        private List<Enquiry> enquiry_list = new List<Enquiry>();
        public void MainProgram(Pragma p)
        {
            
            enquiry_list.Add(new Enquiry { Name = "Galileo Galilei", Email = "n/a", Phone = "n/a", Acc = p.BigString });
            enquiry_list.Add(new Enquiry { Name = "Patrick Moore", Email = "pmoore@baa.co.uk", Phone = "00449881234", Acc = p.BigString });
            enquiry_list.Add(new Enquiry { Name = "Carl Sagan", Email = "carl@uh.com", Phone = "19883456", Acc = p.BigString });
    
        }

        public void BuildList(Pragma p) { MainProgram(p); }

        public IEnumerator GetEnumerator()
        {
            return enquiry_list.GetEnumerator();
        }
    }
    public interface IEnumerableClass : IEnumerable
    {
        new IEnumerator GetEnumerator();
    }
    //test class for VBA interop
    [ClassInterface(ClassInterfaceType.None)]
    public class MyEnumerableClass : IEnumerableClass
    {
        public IEnumerator GetEnumerator()
        {
            List<Enquiry> enquiry_list = new List<Enquiry>();
            enquiry_list.Add(new Enquiry { Name = "Galileo Galilei", Email = "n/a", Phone = "n/a", Acc = "GetS()" });
            enquiry_list.Add(new Enquiry { Name = "Patrick Moore", Email = "pmoore@baa.co.uk", Phone = "00449881234", Acc= "p.CredentialPath" });
            enquiry_list.Add(new Enquiry { Name = "Carl Sagan", Email = "carl@uh.com", Phone = "19883456", Acc = "p.CredentialPath" });
            
            return enquiry_list.GetEnumerator();
        }
    }

    public interface IEnumerable_ReadEmailClass : IEnumerable
    {
        new IEnumerator GetEnumerator();
        void ReadEmails(ProgramSettings PS);
    }
    public class ReadEmail { 

        static byte[] FromBase64ForUrlString(string base64ForUrlInput)
        {
            int padChars = (base64ForUrlInput.Length % 4) == 0 ? 0 : (4 - (base64ForUrlInput.Length % 4));
            System.Text.StringBuilder thisresult = new System.Text.StringBuilder(base64ForUrlInput, base64ForUrlInput.Length + padChars);
            thisresult.Append(String.Empty.PadRight(padChars, '='));
            thisresult.Replace('-', '+');
            thisresult.Replace('_', '/');
            return Convert.FromBase64String(thisresult.ToString());
        }

        
        static Enquiry ReadEnquiry(string s, string account)
        {
            //myhead a bit messes up - cant think of an easier way to do this.
            //1. remove unneeded text before and after body of enquriy
            //2. convert to array - one element per enq item
            //3. for ea array element - split on ":"
            //4. CASE statement Name:Vlaue and create a new Enquiry obejct.


            //Parse string to Enquiry Object
            Enquiry e = new Enquiry();
            
            if (s.Contains("QuickBus.ie"))
            {
                e.Source = "QuickBus";
                e.EnquiryDateFormat = "dd/MM/yyyy";
            }
            else if (s.Contains("Fastbus.ie"))
            {
                e.Source = "FastBus";
                e.EnquiryDateFormat = "MM/dd/yyyy";
            }
            else if (s.Contains("LocalLink"))
            {
                e.Source = "LocalLink";
                e.EnquiryDateFormat = "dd/MM/yyyy";
            }
            else
            {
                e.Source = "Unknown";
                e.EnquiryDateFormat = "dd/MM/yyyy";
            }

            e.Acc = account;

            //substring between ie and Map
            int pFrom = s.IndexOf(".ie") + ".ie".Length;
            int pTo = s.LastIndexOf("Map");
            s = s.Substring(pFrom, pTo - pFrom).Trim();

            //Console.WriteLine(s);

            string[] stringSeparators = new string[] { "\r\n" };
            string[] EnquiryItems = s.Split(stringSeparators, StringSplitOptions.None); //each line as an element in the array "Name: Jean-Luc Picard"

            string myDate = "1/1/1900";
            string myTime = "1/1/1900";

            foreach (var EnquiryItem in EnquiryItems)
            {

                string[] NameValues = EnquiryItem.Split(':');
                //foreach (string nv in NameValues) { Console.Write(nv); }
                switch (NameValues[0])
                {
                    case "Name":
                        e.Name = NameValues[1].Trim();
                        break;
                    case "Phone":
                        e.Phone = NameValues[1].Trim(); ;
                        break;
                    case "Email":
                        e.Email = NameValues[1].Trim(); ;
                        break;
                    case "Bus":
                        e.Bus = NameValues[1].Trim();
                        break;
                    case "Selected Bus":
                        e.Bus = NameValues[1].Trim();
                        break;
                    case "Pickup":
                        e.Pickup = NameValues[1].Trim();
                        break;
                    case "Drop off":
                        e.Dest = NameValues[1].Trim();
                        break;
                    case "Destination":
                        e.Dest = NameValues[1].Trim();
                        break;
                    case "Return Trip":
                        e.Return = NameValues[1].Trim();
                        break;
                    case "Pick-up Date":
                        myDate = NameValues[1].TrimStart();
                        break;
                    case "Pick-up Time":
                        myTime = NameValues[1].Trim();
                        break;
                    case "Pickup Date/Time":
                        //string format=  Pickup Date/Time: dd/mm/yy hh:mm
                        //NameValues[0] = field name 
                        //NameValues[1] = date plus hour part 
                        //NameValues[2] = minute part
                        //remove leading space, split on whitespsace to get datepart & hourpart, concat hourpart with minutes
                        // 
                        myDate = NameValues[1].TrimStart();
                        string[] dateparts = NameValues[1].TrimStart().Split((char)32);
                        myDate = dateparts[0];
                        myTime = dateparts[1] + ":" + NameValues[2];
                        break;
                }
            }         

            if (myDate != "1/1/1900")
            {
                DateTime thisDate;

      
                if (e.EnquiryDateFormat == "dd/MM/yyyy") {
                    thisDate = DateTime.Parse(myDate.ToString());
                }
                else {
                    if (!(DateTime.TryParseExact(myDate, e.EnquiryDateFormat, CultureInfo.InvariantCulture,
                          DateTimeStyles.None, out thisDate)))
                    {
                        thisDate = DateTime.Parse(myDate.ToString()); 
                    }
                }
                e.PickDate = thisDate;
            }
            if (myTime != "1/1/1900")
            {

                //DateTime nTime =  DateTime.Parse(myTime + ":00", new CultureInfo("ie"));
                DateTime t = DateTime.Parse(myTime + ":00", new CultureInfo("ie"));
                DateTime d = e.PickDate;
         
                e.PickDate = new DateTime(d.Year, d.Month, d.Day, t.Hour, t.Minute, t.Second);
            }

            return e;
        }

        public void MoveTokenFile() {
            
            foreach (string newPath in Directory.GetFiles("token", "*.json*", SearchOption.TopDirectoryOnly))
            {
                File.Copy(newPath, "C:\\Projects\\ReadEmails\\ReadEmails_Caller\\bin\\Debug\\token.json", true);
            }
        }

        //[DispId(-4)]
        //gets popuated by ReadEmails
        //to get this to work for com interop with access vba 
        //I have created the ListofEnquiries collection outside the main(ReadEmails) method.
        //I can now use GetEnumerator to return a collection of Enquiry objects that vba can loop over.
        public List<Enquiry> ListofEnquiries;

        public IEnumerator GetEnumerator()
        {
            return ListofEnquiries.GetEnumerator();
        }
        // the vba caller will 
        // 1. create a new instanse of ProgramSettings and populate accordingly
        // 2. call the ReadEmails method and populate the collection
        // 3. use a for each loop to iterate over the collection

        //public List<Enquiry> ReadEmails(ProgramSettings PS)
        public void ReadEmails(ProgramSettings PS)
        {
            // If modifying these scopes, delete your previously saved credentials
            // at ~/.credentials/gmail-dotnet-quickstart.json
            string[] Scopes = { GmailService.Scope.GmailReadonly };
            string ApplicationName = "Gmail API .NET Read Form Enquiry Data";
           
            UserCredential credential;
            
            using (var stream =
               new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                //string credPath = PS.CredentialPath;
                string credPath = PS.CredentialPath;

                if (PS.ClearAccessToken & Directory.Exists(credPath))
                {
                    Directory.Delete(credPath, true);
                }

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
                ;
            }
            // Create Gmail API service.
            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            List<Message> result = new List<Message>();
            UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List("me");

            UsersResource.GetProfileRequest profile_request = service.Users.GetProfile("me");
            Profile profile_response = profile_request.Execute();
            string acc = profile_response.EmailAddress;

            request.IncludeSpamTrash = PS.IncSpamTrash;
            request.LabelIds = "INBOX";
            request.MaxResults = PS.NoOfEmailsToRead;
            string query = "";
            string query_from = ""; //use to filter for an address
            query = "label:unread";
            query = query + " " + "subject:({quickbus.ie fastbus.ie})";
            query = query + " " + query_from;

            request.Q = query;
            ListMessagesResponse response = request.Execute();

            List<Enquiry> Enquiries = new List<Enquiry>();
   
            if (response != null && response.Messages != null)
            {
                foreach (Message m in response.Messages)
                {
                    Message ThisEmail = service.Users.Messages.Get("me", m.Id).Execute();
                    if (ThisEmail != null)
                    {
                        String from = "";
                        String date = "";
                        String subject = "";
                        String messageID = "";
                        String decodedString = "";

                        foreach (MessagePartHeader headerpart in ThisEmail.Payload.Headers)
                        {
                            if (headerpart.Name == "Date")
                            {
                                date = headerpart.Value;
                            }
                            else if (headerpart.Name == "From")
                            {
                                from = headerpart.Value;
                            }
                            else if (headerpart.Name == "Subject")
                            {
                                subject = headerpart.Value;
                            }
                            else if (headerpart.Name == "Message-ID")
                            {
                                messageID = headerpart.Value;
                            }


                            if (date != "" && from != "")
                            {
                                foreach (MessagePart p in ThisEmail.Payload.Parts)
                                {
                                    //Console.WriteLine(p.MimeType);

                                    if ((p.MimeType == "text/plain"))

                                    {
                                        byte[] data = FromBase64ForUrlString(p.Body.Data);
                                        decodedString = System.Text.Encoding.UTF8.GetString(data);
                                    }
                                }
                            }
                        }
                        //Console.WriteLine(decodedString);
                        
                        Enquiry ThisEmailEnquiry = ReadEnquiry(decodedString, acc);
                        Enquiries.Add(ThisEmailEnquiry);
                    }

                }
            }
            //return Enquiries;
            ListofEnquiries = Enquiries;
        }
    }
}
