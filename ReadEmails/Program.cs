//14072021
//jkelly@jsd.ie
//Jonathan Kelly
//coder187@hotmail.com
//program to read email or emails from given gmail inbox or email id and then 
//perist to db or otherwise make data available to caller...

//https://developers.google.com/gmail/api/quickstart/dotnet


using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Text;

namespace ReadEmails
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/gmail-dotnet-quickstart.json
        static string[] Scopes = { GmailService.Scope.GmailReadonly };
        static string ApplicationName = "Gmail API .NET Read Form Enquiry Data";
        static void Main(string[] args)
        {
            UserCredential credential;
            
            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
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

            //SAMPLE CODE
            // Define parameters of request.
            //UsersResource.LabelsResource.ListRequest request = service.Users.Labels.List("me");
            //UsersResource.MessagesResource.ListRequest request2 = service.Users.Messages.List("me");

            //// List labels.
            //IList<Label> labels = request.Execute().Labels;
            //Console.WriteLine("Labels:");
            //if (labels != null && labels.Count > 0)
            //{
            //    foreach (var labelItem in labels)
            //    {
            //        Console.WriteLine("{0}", labelItem.Name);
            //    }
            //}
            //else
            //{
            //    Console.WriteLine("No labels found.");
            //}
            ////Console.Read();

            //// List emails.
            //IList<Message> emails = request2.Execute().;
            //long emailSizeEst = (long)request2.Execute().ResultSizeEstimate;

            //Console.WriteLine("Emails:");
            //Console.WriteLine("Email Size:" + emailSizeEst);

            //if (emails != null && emails.Count > 0)
            //{
            //    foreach (var emailItem in emails)
            //    {
            //        Console.WriteLine("{0}", emailItem.Id);
            //    }
            //}
            //else
            //{
            //    Console.WriteLine("No emails found.");
            //}
            

            long pageCount = 0;
            List<Message> result = new List<Message>();
            UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List("me");

            UsersResource.GetProfileRequest profile_request = service.Users.GetProfile("me");
            Profile profile_response = profile_request.Execute();
            string acc = profile_response.EmailAddress;
            

            request.IncludeSpamTrash = false;
            request.LabelIds= "INBOX";
            request.MaxResults = 10;
            string query = "";
            query = "label:unread";
            query = query + " " + "subject:({quickbus.ie fastbus.ie})";

            request.Q = query;
            List<Enquiry> Enquiries = new List<Enquiry>();

            ListMessagesResponse response = request.Execute();
            if (response !=null && response.Messages !=null) {
                foreach (Message m  in response.Messages) {
                    Message ThisEmail = service.Users.Messages.Get("me", m.Id).Execute();
                    if (ThisEmail != null) {
                        //Console.WriteLine(ThisEmail.Snippet);
                        //Console.WriteLine(ThisEmail.Id);

                        

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
                            } else if (headerpart.Name == "Message-ID") 
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
                                        //Console.WriteLine(p.MimeType);
                                        byte[] data = FromBase64ForUrlString(p.Body.Data);
                                        decodedString = System.Text.Encoding.UTF8.GetString(data);
                                        //Console.WriteLine(decodedString);


                                    }
                                    
                                    //Console.WriteLine("############################################################");
                                }

                            }
                            
                        }
                        //Console.WriteLine(decodedString);
                        Enquiry ThisEmailEnquiry = ReadEnquiry(decodedString,acc);
                        Enquiries.Add(ThisEmailEnquiry);
                    }

                }
            }

            foreach (Enquiry e in Enquiries){
                WriteEnquiry(e);
            }
            Console.Read();



            byte[] FromBase64ForUrlString(string base64ForUrlInput)
                {
                    int padChars = (base64ForUrlInput.Length % 4) == 0 ? 0 : (4 - (base64ForUrlInput.Length % 4));
                    System.Text.StringBuilder thisresult = new System.Text.StringBuilder(base64ForUrlInput, base64ForUrlInput.Length + padChars);
                    thisresult.Append(String.Empty.PadRight(padChars, '='));
                    thisresult.Replace('-', '+');
                    thisresult.Replace('_', '/');
                    return Convert.FromBase64String(thisresult.ToString());
                }
                
            string Base64UrlEncode(string input)
            {
                var inputBytes = System.Text.Encoding.UTF8.GetBytes(input);
                return Convert.ToBase64String(inputBytes).Replace("+", "-").Replace("/", "_").Replace("=", "");
            }

            Enquiry ReadEnquiry(string s, string account)
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
                            e.PickDate = NameValues[1].Trim();
                            break;
                    }
                }
                return e;
            }
               
            void WriteEnquiry(Enquiry e) {
                Console.WriteLine("//////////////////////////////////////////////////");
                Console.WriteLine(e.Acc);
                Console.WriteLine(e.Source);
                Console.WriteLine(e.Name);
                Console.WriteLine(e.Phone);
                Console.WriteLine(e.Email);
                Console.WriteLine(e.Bus);
                Console.WriteLine(e.Pickup);
                Console.WriteLine(e.Dest);
                Console.WriteLine(e.PickDate);
                Console.WriteLine(e.Return);
                Console.WriteLine("//////////////////////////////////////////////////");
            }
        }
    }
}

class Enquiry
{
    public string Name { get; set; }
    public string Email { get; set; }
    public string Phone { get; set; }
    public string Bus { get; set; }
    public string Pickup { get; set; }
    public string PickDate { get; set; }
    public string Dest { get; set; }
    public string Return { get; set; }
    public string Source { get; set; }
    public string Acc { get; set; }


}
