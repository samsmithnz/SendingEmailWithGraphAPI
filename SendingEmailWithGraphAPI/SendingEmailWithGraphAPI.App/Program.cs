using Microsoft.Extensions.Configuration;
using System;
using System.Threading.Tasks;

namespace GraphEmail
{
    class Program
    {       
        static async Task Main(string[] args)
        {
            //Load configurations
            IConfigurationBuilder config = new ConfigurationBuilder()
               .AddUserSecrets<Program>();
            IConfigurationRoot Configuration = config.Build();
            string tenantId = Configuration["TenantId"];
            string clientId = Configuration["ClientId"];
            string clientSecret = Configuration["ClientSecret"]; 
            string userId = Configuration["UserId"];

            //Build attachments
            string[] attachments = new string[2];
            attachments[0] = @"C:\Users\samsmit\source\repos\SendingEmailWithGraphAPI\SendingEmailWithGraphAPI\SendingEmailWithGraphAPI.App\Images\Library.jpg";
            attachments[1] = @"C:\Users\samsmit\source\repos\SendingEmailWithGraphAPI\SendingEmailWithGraphAPI\SendingEmailWithGraphAPI.App\Images\Hyperspace.jpg";

            //Send the email
            GraphHelper graphHelper = new GraphHelper(tenantId, clientId, clientSecret, userId);
            await graphHelper.SendEmail("samsmithmsft@ssnzmsft.onmicrosoft.com", "samsmithnz@gmail.com", "email test " + DateTime.Now.ToString(), "Hello world", attachments);
        }
    }
}
