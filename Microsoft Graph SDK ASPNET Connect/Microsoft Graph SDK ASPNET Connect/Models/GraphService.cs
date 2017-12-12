/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

using Microsoft.Graph;
using Resources;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft_Graph_SDK_ASPNET_Connect.Helpers;
using Newtonsoft.Json.Linq;

namespace Microsoft_Graph_SDK_ASPNET_Connect.Models
{
    public class GraphService
    {

        public async Task<string> GetAllChannels(GraphServiceClient graphClient)
        {
            var token =  await SampleAuthProvider.Instance.GetUserAccessTokenAsync();

            var groups = await graphClient.Groups.Request().GetAsync();
            foreach (var group in groups)
            {
                //var conversations = group.Conversations;
                //foreach (var conversation in conversations)
                //{
                //    ;
                //}
                var client = new HttpClient();
                //var client = new WebClient();
                //var requestBody = new Dictionary<string, string>()
                //{
                //    {"client_id", ConfigurationManager.AppSettings["MicrosoftAppId"] },
                //    {"grant_type", "client_credentials"},
                //    { "scope", ConfigurationManager.AppSettings["Scope"]},
                //    {"client_secret", ConfigurationManager.AppSettings["MicrosoftAppPassword"]}
                //};
                //client.Headers.Add("Authorization", "Bearer " + token);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                var url = String.Format("https://graph.microsoft.com/beta/groups/" + group.Id + "/channels?$format=json");

                //var content = new FormUrlEncodedContent(requestBody);
                //eyJ0eXAiOiJKV1QiLCJub25jZSI6IkFRQUJBQUFBQUFCSGg0a21TX2FLVDVYcmp6eFJBdEh6a3pFQXI3UWdFdzFjTlQzeFFUWnpYZEtnQk1STzZKMnJXaE1VV3huZGFzbFBjWTJoMnltTVdlZ2t3WENMUHNDSmVTWU1rMGlfblAyTmR2WnJLTGJTOXlBQSIsImFsZyI6IlJTMjU2IiwieDV0IjoieDQ3OHh5T3Bsc00xSDdOWGs3U3gxN3gxdXBjIiwia2lkIjoieDQ3OHh5T3Bsc00xSDdOWGs3U3gxN3gxdXBjIn0.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9mYjVhNWZjMi1iOWZlLTRmNzQtYWFkNS01YjdhNjU2OTVkYTYvIiwiaWF0IjoxNTEzMDQxNTAzLCJuYmYiOjE1MTMwNDE1MDMsImV4cCI6MTUxMzA0NTQwMywiYWNyIjoiMSIsImFpbyI6IkFTUUEyLzhHQUFBQVYvdjNBOGg5cFFqMWoyb0lCbHFpNXRFZnhZNTQ4c3k2cmtkWDgvZlZ2eVU9IiwiYW1yIjpbInB3ZCJdLCJhcHBfZGlzcGxheW5hbWUiOiJHcmFwaCBTREsgQVNQTkVUIENvbm5lY3QiLCJhcHBpZCI6Ijc2MWI1MDFmLTQxNGEtNDg4MS1hYWRhLTFmOGYxOTkwOTI4NCIsImFwcGlkYWNyIjoiMSIsImVfZXhwIjoyNjI4MDAsImZhbWlseV9uYW1lIjoiQmxvY2giLCJnaXZlbl9uYW1lIjoiTGF1cmEiLCJpcGFkZHIiOiIxNzQuMTAzLjExNi4xNDYiLCJuYW1lIjoiTGF1cmEgQmxvY2giLCJvaWQiOiIzYWM0MWM3OC0xNWM5LTRlOTQtOWIxYy1jYzMwZDkzMWZjNTIiLCJwbGF0ZiI6IjMiLCJwdWlkIjoiMTAwMzAwMDBBNjgzRkU3MSIsInNjcCI6IkZpbGVzLlJlYWRXcml0ZSBHcm91cC5SZWFkLkFsbCBNYWlsLlNlbmQgVXNlci5SZWFkIiwic3ViIjoiTHVIOXZjNkk1QXhwLWZmWnpPWTRrdVppNU5hdGVveWdncnBGMjBJNkM4ZyIsInRpZCI6ImZiNWE1ZmMyLWI5ZmUtNGY3NC1hYWQ1LTViN2E2NTY5NWRhNiIsInVuaXF1ZV9uYW1lIjoibGJAbGF1cmFibG9jaC5vbm1pY3Jvc29mdC5jb20iLCJ1cG4iOiJsYkBsYXVyYWJsb2NoLm9ubWljcm9zb2Z0LmNvbSIsInV0aSI6IkM4Q2RZNzJBa2tlR1owX1ZPMXczQUEiLCJ2ZXIiOiIxLjAiLCJ3aWRzIjpbIjYyZTkwMzk0LTY5ZjUtNDIzNy05MTkwLTAxMjE3NzE0NWUxMCJdfQ.Z15KiHstaRS8PVmEj8TvbUgVMHEdKh8WRuAT_uelcNmUcVG1eZZhhvaae75-CEToU065XyNcCCCbUWbfL-H3xMEC6zHuMv5Uy_L6P1BIvjcvUKU42eSve1AaXu2hoEEk6-HzxOhlQQd8mT3UBe6icL4O8HuxKl6_MmwDX7v6umDYIau9bgzs7HZAwlduI1QmUI8qpX5zPUh81t_aALWnnE8KVC_oMoe-3hijRF-DLUVYVACxaEeWRIKD_AenCpFuc23Qy5NZAgOd72-q32fu7dUg69YJ6OT0P4-7eufVLBvH-DXztkfdVoJhuCZAizEpRdHPlGkkne08ee9cB6Ub9A
                //TODO SEND AUTH TOKEN IN HEADER
                HttpResponseMessage response = await client.GetAsync(url);
                if (response.IsSuccessStatusCode)
                {
                    var getResponse = await response.Content.ReadAsStringAsync();
                    var responseObject = JObject.Parse(getResponse);
                    var channel = responseObject["id"];
                    return channel.ToString();
                }

            }
            return null;
        }

        // Get the current user's email address from their profile.
        public async Task<string> GetMyEmailAddress(GraphServiceClient graphClient)
        {

            // Get the current user. 
            // This sample only needs the user's email address, so select the mail and userPrincipalName properties.
            // If the mail property isn't defined, userPrincipalName should map to the email for all account types. 
            User me = await graphClient.Me.Request().Select("mail,userPrincipalName").GetAsync();
            return me.Mail ?? me.UserPrincipalName;
        }

        // Send an email message from the current user.
        public async Task SendEmail(GraphServiceClient graphClient, Message message)
        {
            await graphClient.Me.SendMail(message, true).Request().PostAsync();
        }

        // Create the email message.
        public async Task<Message> BuildEmailMessage(GraphServiceClient graphClient, string recipients, string subject)
        {

            // Get current user photo
            Stream photoStream = await GetCurrentUserPhotoStreamAsync(graphClient);


            // If the user doesn't have a photo, or if the user account is MSA, we use a default photo

            if ( photoStream == null)
            {
                photoStream = System.IO.File.OpenRead(System.Web.Hosting.HostingEnvironment.MapPath("/Content/test.jpg"));
            }

            MemoryStream photoStreamMS = new MemoryStream();
            // Copy stream to MemoryStream object so that it can be converted to byte array.
            photoStream.CopyTo(photoStreamMS);

            DriveItem photoFile = await UploadFileToOneDrive(graphClient, photoStreamMS.ToArray());

            MessageAttachmentsCollectionPage attachments = new MessageAttachmentsCollectionPage();
            attachments.Add(new FileAttachment
            {
                ODataType = "#microsoft.graph.fileAttachment",
                ContentBytes = photoStreamMS.ToArray(),
                ContentType = "image/png",
                Name = "me.png"
            });

            Permission sharingLink = await GetSharingLinkAsync(graphClient, photoFile.Id);

            // Add the sharing link to the email body.
            string bodyContent = string.Format(Resource.Graph_SendMail_Body_Content, sharingLink.Link.WebUrl);

            // Prepare the recipient list.
            string[] splitter = { ";" };
            string[] splitRecipientsString = recipients.Split(splitter, StringSplitOptions.RemoveEmptyEntries);
            List<Recipient> recipientList = new List<Recipient>();
            foreach (string recipient in splitRecipientsString)
            {
                recipientList.Add(new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = recipient.Trim()
                    }
                });
            }

            // Build the email message.
            Message email = new Message
            {
                Body = new ItemBody
                {
                    Content = bodyContent,
                    ContentType = BodyType.Html,
                },
                Subject = subject,
                ToRecipients = recipientList,
                Attachments = attachments
            };
            return email;
        }

        // Gets the stream content of the signed-in user's photo. 
        // This snippet doesn't work with consumer accounts.
        public async Task<Stream> GetCurrentUserPhotoStreamAsync(GraphServiceClient graphClient)
        {
            Stream currentUserPhotoStream = null;

            try
            {
                currentUserPhotoStream = await graphClient.Me.Photo.Content.Request().GetAsync();

            }

            // If the user account is MSA (not work or school), the service will throw an exception.
            catch (ServiceException)
            {
                return null;
            }

            return currentUserPhotoStream;

        }

        // Uploads the specified file to the user's root OneDrive directory.
        public async Task<DriveItem> UploadFileToOneDrive(GraphServiceClient graphClient, byte[] file)
        {
            DriveItem uploadedFile = null;

            try
            {
                MemoryStream fileStream = new MemoryStream(file);
                uploadedFile = await graphClient.Me.Drive.Root.ItemWithPath("me.png").Content.Request().PutAsync<DriveItem>(fileStream);

            }


            catch (ServiceException)
            {
                return null;
            }

            return uploadedFile;
        }

        public static async Task<Permission> GetSharingLinkAsync(GraphServiceClient graphClient, string Id)
        {
            Permission permission = null;

            try
            {
                permission = await graphClient.Me.Drive.Items[Id].CreateLink("view").Request().PostAsync();
            }

            catch (ServiceException)
            {
                return null;
            }

            return permission;
        }

    }
}