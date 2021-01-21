using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace GraphEmail
{
    //This link was most helpful (open the tabs in private windows if Medium gives you grief): 
    //https://jatindersingh81.medium.com/c-code-to-to-send-emails-using-microsoft-graph-api-2a90da6d648a
    public class GraphHelper
    {
        private readonly string _tenantId = "";
        private readonly string _clientId = "";
        private readonly string _clientSecret = "";
        private readonly string _userId;

        //The following scope is required to acquire the token
        private readonly string[] _scopes = new string[] { "https://graph.microsoft.com/.default" };

        public GraphHelper(string tenantId, string clientId, string clientSecret, string userId)
        {
            _tenantId = tenantId;
            _clientId = clientId;
            _clientSecret = clientSecret;
            _userId = userId;
        }

        //Create the email message
        public Message CreateEmail(string from, string to, string subject, string htmlBody,
            string[] attachments = null, string cc = null, string bcc = null)
        {
            Message message = new Message
            {
                From = new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = from
                    }
                },
                ToRecipients = BuildEmailRecipientList(to),
                Subject = subject,
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = htmlBody
                },
                Attachments = BuildAttachments(attachments)
            };
            if (string.IsNullOrEmpty(cc) == false)
            {
                message.CcRecipients = BuildEmailRecipientList(cc);
            }
            if (string.IsNullOrEmpty(bcc) == false)
            {
                message.CcRecipients = BuildEmailRecipientList(bcc);
            }

            return message;
        }

        //Create the message, authenicate and send the message
        public async Task SendEmail(string from, string to, string subject, string htmlBody,
            string[] attachments = null, string cc = null, string bcc = null)
        {
            //Create the message
            Message message = CreateEmail(from, to, subject, htmlBody, attachments, cc, bcc);

            //Authenicate
            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
                .Create(_clientId)
                .WithTenantId(_tenantId)
                .WithClientSecret(_clientSecret)
                .Build();

            AuthenticationResult authResultDirect = await confidentialClientApplication
                .AcquireTokenForClient(_scopes)
                .ExecuteAsync()
                .ConfigureAwait(false);

            ClientCredentialProvider authProvider = new ClientCredentialProvider(confidentialClientApplication); //Microsoft.Graph.Auth is required for the following to work
            GraphServiceClient graphClient = new GraphServiceClient(authProvider);

            //Send the message
            await graphClient.Users[_userId]
                    .SendMail(message, false)
                    .Request()
                    .PostAsync();
        }

        //Upload the attachments
        private MessageAttachmentsCollectionPage BuildAttachments(string[] files)
        {
            MessageAttachmentsCollectionPage attachments = null;
            if (files != null && files.Length > 0)
            {
                attachments = new MessageAttachmentsCollectionPage();

                foreach (string file in files)
                {
                    // Create the message with attachment.
                    string fileName = Path.GetFileName(file);
                    string fileNameWithNoExtension = Path.GetFileName(file).Replace(Path.GetExtension(fileName), "");
                    byte[] contentBytes = System.IO.File.ReadAllBytes(file);
                    string contentType = "image/jpg";
                    attachments.Add(new FileAttachment
                    {
                        ODataType = "#microsoft.graph.fileAttachment",
                        ContentBytes = contentBytes,
                        ContentType = contentType,
                        ContentId = fileNameWithNoExtension,
                        Name = fileName
                    });
                }

            }
            return attachments;
        }

        //Look at a recipient list and split by , to create a list of email addresses to send/cc/bcc to
        private List<Recipient> BuildEmailRecipientList(string emailAddresses)
        {
            string[] toList = emailAddresses.Split(','); // big assumption here that they are separated by ",". May need to refactor to also include ;
            List<Recipient> emails = new List<Recipient>();
            foreach (string email in toList)
            {
                emails.Add(new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = email
                    }
                });
            }
            return emails;
        }

    }
}
