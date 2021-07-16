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
            request.IncludeSpamTrash = false;
            request.LabelIds= "INBOX";
            request.MaxResults = 1;
            //request.Q = query;
            
            ListMessagesResponse response = request.Execute();
            if (response !=null && response.Messages !=null) {
                foreach (Message m  in response.Messages) {
                    Message ThisEmail = service.Users.Messages.Get("me", m.Id).Execute();
                    if (ThisEmail != null) {
                        Console.WriteLine(ThisEmail.Snippet);
                        Console.WriteLine(ThisEmail.Id);

                        

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
                                        Console.WriteLine(p.MimeType);
                                        byte[] data = FromBase64ForUrlString(p.Body.Data);
                                        decodedString = System.Text.Encoding.UTF8.GetString(data);
                                        Console.WriteLine(decodedString);


                                    }
                                    
                                    //Console.WriteLine("############################################################");
                                }

                            }
                            
                        }
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine(date + "\r\n" + from + "\r\n" + subject + "\r\n" + messageID + "\r\n" + decodedString);
                    }

                }
            }
            //do
            //{
            //  try
            //{


            //result.AddRange(response.Messages);
            //        pageCount++;
            //        Console.WriteLine(pageCount);
            ////request.PageToken = response.NextPageToken;
            ////}
            ////catch (Exception e)
            // {
            //Console.WriteLine("An error occurred: " + e.Message);
            // }

            //} while (!String.IsNullOrEmpty(request.PageToken));

            //Console.WriteLine("Found " + pageCount + " messages.");

            //int i = 1;
            //foreach (Message m in result)
            //{
            //    Console.WriteLine(m.Id);

            //    Console.WriteLine(Base64UrlEncode(m.Raw));

            //    Console.WriteLine(i + "###############################################################");
            //    i++;
            //}
            
                
               
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
        }
        
    }

    
}
