using System;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using System.Security;

namespace SP_CSOM_O365_Alerts
{
    class Program
    {
        private static ConsoleColor defaultForeground;
        static void Main(string[] args)
        {
            defaultForeground = Console.ForegroundColor;

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Enter the URL to your Office 365 SharePoint site:");

            Console.ForegroundColor = defaultForeground;
            string webUrl = Console.ReadLine();
            //string webUrl = "https://mytenant.sharepoint.com";

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Enter your user name (ex: user@mytenant.onmicrosoft.com):");
            Console.ForegroundColor = defaultForeground;
            string userName = Console.ReadLine();
            //string userName = "user@mytenant.onmicrosoft.com";

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Enter your password:");
            Console.ForegroundColor = defaultForeground;
            SecureString password = CreateSecureStringPasswordFromConsoleInput();

            using (var context = new ClientContext(webUrl))
            {
                context.Credentials = new SharePointOnlineCredentials(userName, password);
                context.Load(context.Web, w => w.Title);
                context.ExecuteQuery();

                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("Connected to site: " + context.Web.Title);
                Console.WriteLine("");
                Console.ForegroundColor = defaultForeground;

                CreateMyAlert(context);

                GetMyAlerts(context);
            }

            Console.ForegroundColor = defaultForeground;
            Console.WriteLine("Press a key to exit.");
            Console.ReadKey();
        }

        private static void GetMyAlerts(ClientContext context)
        {
            Web web = context.Web;
            context.Load(web);
            context.Load(web.Lists);

            User currentUser = context.Web.CurrentUser;
            context.Load(currentUser);
            context.Load(currentUser.Alerts);
            context.Load(currentUser.Alerts,
                         lists => lists.Include(
                             list => list.Title,
                                    list => list.ListID));

            AlertCollection myAlerts = currentUser.Alerts;

            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("Getting alerts for user " + currentUser.Title);
            Console.WriteLine("");

            context.ExecuteQuery();

            Console.WriteLine(myAlerts.Count + " alerts found for user " + currentUser.Title);
            Console.WriteLine("");

            foreach (Alert alert in myAlerts)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("\\---------[Alert Metadata]---------/");
                Console.ForegroundColor = defaultForeground;
                Console.WriteLine("ID: " + alert.ID);
                Console.WriteLine("Item: " + alert.Item);
                Console.WriteLine("Title: " + alert.Title);
                Console.WriteLine("AlertFrequency: " + alert.AlertFrequency);
                Console.WriteLine("AlertTemplateName: " + alert.AlertTemplateName);
                Console.WriteLine("AlertType: " + alert.AlertType);
                Console.WriteLine("AlwaysNotify: " + alert.AlwaysNotify);
                Console.WriteLine("DeliveryChannels: " + alert.DeliveryChannels);
                Console.WriteLine("EventType: " + alert.EventType);
                Console.WriteLine("Status: " + alert.Status);
                Console.WriteLine("ListID: " + alert.ListID);
                Console.WriteLine("");
            }
        }

        private static void CreateMyAlert(ClientContext context)
        {
            Console.ForegroundColor = ConsoleColor.White;

            Web web = context.Web;
            context.Load(web);
            context.Load(web.Lists);

            User currentUser = context.Web.CurrentUser;
            context.Load(currentUser);
            context.Load(currentUser.Alerts);
            context.Load(currentUser.Alerts,
                         lists => lists.Include(
                             list => list.Title,
                                    list => list.ListID));

            context.ExecuteQuery();

            AlertCreationInformation myNewAlert = new AlertCreationInformation();
            myNewAlert.List = context.Web.Lists.GetByTitle("Documents");
            myNewAlert.AlertFrequency = AlertFrequency.Daily;
            myNewAlert.AlertTime = DateTime.Today.AddDays(1);
            myNewAlert.AlertType = AlertType.List;
            myNewAlert.AlwaysNotify = false;
            myNewAlert.DeliveryChannels = AlertDeliveryChannel.Email;

            myNewAlert.Status = AlertStatus.On;
            myNewAlert.Title = "My new alert created at : " + DateTime.Now.ToString();
            myNewAlert.User = currentUser;

            // These two properties currently have no impact on alert creation.
            // This is a known bug.  If you set these properties and do not use the property bag
            // you will receive an Object reference not set to an instance of an object exception.
            // myNewAlert.EventType = AlertEventType.All;
            // myNewAlert.Filter = "0";

            // Currently, you need to use the property bag entries below to create the alert.
            Dictionary<string, string> properties = new Dictionary<string, string>()
            {
                { "eventtypeindex", "0" },
                    //Change Type: 
                    // 0 = All Changes
                    // 1 = New items added
                    // 2 = New items are added
                    // 3 = Existing items are modified 
                { "FilterIndex", "0" } 
                    //Send Me and alert when: 
                    // 0 = Anything Changes
                    // 1 = Someone else changes a document
                    // 2 = Someone else changes a document created by me
                    // 3 = Someone else changes a document modified by me
            };

            // You can also set addition properties to configure the alert, such as:
            // properties.Add("dispformurl", "Shared Documents/Forms/DispForm.aspx");
            // properties.Add("defaultitemopen", "Browser");
            // properties.Add("sendurlinsms", "False");
            // properties.Add("mobileurl", "https://cand3.sharepoint.com/_layouts/15/mobile/");
            // properties.Add("siteurl", "https://cand3.sharepoint.com");

            myNewAlert.Properties = properties;

            var newAlertGuid = currentUser.Alerts.Add(myNewAlert);

            currentUser.Update();

            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("Creating alert for user " + currentUser.Title);
            Console.WriteLine("");

            context.ExecuteQuery();

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("New alert created with ID = " + newAlertGuid.Value.ToString());
            Console.WriteLine("");
        }

        private static SecureString CreateSecureStringPasswordFromConsoleInput()
        {
            ConsoleKeyInfo info;

            //Get the user's password as a SecureString
            SecureString securePassword = new SecureString();
            do
            {
                info = Console.ReadKey(true);
                if (info.Key != ConsoleKey.Enter)
                {
                    securePassword.AppendChar(info.KeyChar);
                }
            }
            while (info.Key != ConsoleKey.Enter);
            return securePassword;
        }
    }
}
