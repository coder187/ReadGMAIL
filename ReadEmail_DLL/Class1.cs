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
    public class ReadEmail
    {
        

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

            //substring between ie and Map
            int pFrom = s.IndexOf(".ie") + ".ie".Length;
            int pTo = s.LastIndexOf("Map");
            s = s.Substring(pFrom, pTo - pFrom).Trim();

            //Console.WriteLine(s);

            string[] stringSeparators = new string[] { "\r\n" };
            string[] EnquiryItems = s.Split(stringSeparators, StringSplitOptions.None); //each line as an element in the array "Name: Jean-Luc Picard"

            //Parse string to Enquiry Object
            Enquiry e = new Enquiry();
            e.Acc = account;
            e.Source = "FASTBUS/QUICKBUS";

            string myDate = "1/1/1900";
            string myTime = "1/1/1900";

            foreach (var EnquiryItem in EnquiryItems)
            {

                string[] NameValues = EnquiryItem.Split(':');

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
                    case "Pickup":
                        e.Pickup = NameValues[1].Trim();
                        break;
                    case "Drop Off":
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

            //getting errors from DateTime.Parse depending on enquiry source so I'm using
            //TryParse which seems to work.
            if (myDate != "1/1/1900")
            {
                DateTime thisDate;
                if (!(DateTime.TryParse(myDate.ToString(), out thisDate)))
                    thisDate = DateTime.Parse(myDate.ToString(), new CultureInfo("ie"));

                e.PickDate = thisDate;
            }
            if (myTime != "1/1/1900")
            {
                e.PickDate = e.PickDate.Add(TimeSpan.Parse(myTime));
            }

            return e;
        }


       
        public List<Enquiry> ReadEmails(ProgramSettings PS) {
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
            return Enquiries;
        }
    }
}
