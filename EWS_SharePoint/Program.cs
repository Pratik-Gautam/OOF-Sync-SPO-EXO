using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http.Headers;
using System.Web;
using Microsoft.Azure.ActiveDirectory.GraphClient;
using System.Net;
using Microsoft.SharePoint.Client;
using System.Security;

namespace EWS_SharePoint
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine("******************");
                Console.WriteLine("Job Starting......");
                Console.WriteLine("******************");
                Console.ResetColor();

                GetuserEmails();
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine("*************************************************");
                Console.WriteLine("Operation Completed Successfully !");
                Console.WriteLine("*************************************************");
                Console.ReadLine();
                Console.ResetColor();
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Helper method for Exchange Online for rendering autodiscover url
        /// </summary>
        /// <param name="redirectionUrl"></param>
        /// <returns></returns>
        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }

        /// <summary>
        /// Exchange online module to get out of office data and send it to SharePoint Online module
        /// </summary>
        /// <param name="userEmail"></param>
        static void GetOutofOffice(string userEmail,string userDisplayName)
        {

            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            service.Credentials = new WebCredentials(Constants.AdminUsername,Constants.AdminPassword);
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.PrincipalName, userEmail);
            service.AutodiscoverUrl(Constants.AdminUsername, RedirectionUrlValidationCallback);

            OofSettings settings = service.GetUserOofSettings(userEmail);
            Console.ForegroundColor = ConsoleColor.DarkMagenta;
            Console.WriteLine("Retrieved Out of Facility information for "+userEmail+ " from Exchange online successfully.");
            Console.ResetColor();
            if (settings.State==OofState.Enabled)
            {
                //Updating data to SharePoint
                UpdateOOFtoSharePoint("enabled", userDisplayName, null, null);
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("*************************************************");
                Console.WriteLine("User " + userEmail + " is currently Out of office. Updated the details to SharePoint.");
                Console.WriteLine("*************************************************");
                Console.ResetColor();
            }
            else if(settings.State==OofState.Scheduled)
            {
                //Updating data to SharePoint
                UpdateOOFtoSharePoint("scheduled", userDisplayName, settings.Duration.StartTime.ToLongDateString(), settings.Duration.EndTime.ToLongDateString());
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("*************************************************");
                Console.WriteLine("User " + userEmail + " is currently out of office starting from "+settings.Duration.StartTime.ToLongDateString() + " and will return on " + settings.Duration.EndTime.ToLongDateString()+" Updated the details to SharePoint.");
                Console.WriteLine("*************************************************");
                Console.ResetColor();
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("*************************************************");
                Console.WriteLine("User " + userEmail + " is currently in the office");
                Console.WriteLine("*************************************************");
                Console.ResetColor();
            }
        }

        /// <summary>
        /// Azure Active Directory module to fetch user principal and pass to Exchange Online module
        /// </summary>
        static void GetuserEmails()
        {
            #region Setup Active Directory Client

            //*********************************************************************
            // setup Active Directory Client
            //*********************************************************************
            ActiveDirectoryClient activeDirectoryClient;
            try
            {
                activeDirectoryClient = AuthenticationHelper.GetActiveDirectoryClientAsApplication();
            }
            catch (AuthenticationException ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Acquiring a token failed with the following error: {0}", ex.Message);
                if (ex.InnerException != null)
                {
                    //You should implement retry and back-off logic per the guidance given here:http://msdn.microsoft.com/en-us/library/dn168916.aspx
                    //InnerException Message will contain the HTTP error status codes mentioned in the link above
                    Console.WriteLine("Error detail: {0}", ex.InnerException.Message);
                }
                Console.ResetColor();
                Console.ReadKey();
                return;
            }

            #endregion

            #region List of Users by UPN

            //**************************************
            // Demonstrate Getting a list of Users
            //**************************************
            try
            {
                List<IUser> users = activeDirectoryClient.Users.ExecuteAsync().Result.CurrentPage.ToList();
                foreach (IUser user in users)
                {
                    if (!string.IsNullOrEmpty(user.Mail)&&!user.UserPrincipalName.Contains("#EXT#"))
                    { 
                        Console.ForegroundColor = ConsoleColor.DarkYellow;
                        Console.WriteLine("Retrieved user "+ user.UserPrincipalName +" from Azure Active Directory with a valid mailbox successfully.");
                        Console.ResetColor();
                        GetOutofOffice(user.UserPrincipalName,user.DisplayName);                        
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting Users. {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }

            #endregion
        }

        /// <summary>
        /// This method updates OOF information in SharePoint
        /// </summary>
        /// <param name="oofStatus"></param>
        /// <param name="userDisplayName"></param>
        static void UpdateOOFtoSharePoint(string oofStatus,string userDisplayName,string startDate,string endDate)
        {
            //Creating client context for the SharePoint Site
            using (var context = new ClientContext(Constants.SharePointSiteUrl))
            {
                //Creating secure password
                SecureString securePassword = new SecureString();
                foreach (var c in Constants.AdminPassword)
                {
                    securePassword.AppendChar(c);
                }

                context.Credentials = new SharePointOnlineCredentials(Constants.AdminUsername,securePassword);
                List calendarList = context.Web.Lists.GetByTitle("OutOfFacility");
                //Updating metadata when out of facility is scheduled
                if (oofStatus.ToUpper().Equals("SCHEDULED") && !string.IsNullOrEmpty(startDate) && !string.IsNullOrEmpty(endDate))
                {

                    ListItemCreationInformation newItemInfo = new ListItemCreationInformation();
                    ListItem newItem = calendarList.AddItem(newItemInfo);
                    newItem["Title"] = userDisplayName + " is out of Facility";
                    newItem["EventDate"] = startDate;
                    newItem["EndDate"] = endDate;
                    newItem["Category"] = "Out Of Facility";
                    newItem["fAllDayEvent"] = true;
                    newItem.Update();
                }
                //Updating metadata when out of facility is enabled
                else if (oofStatus.ToUpper().Equals("ENABLED") && string.IsNullOrEmpty(startDate) && string.IsNullOrEmpty(endDate))
                {
                    ListItemCreationInformation newItemInfo = new ListItemCreationInformation();
                    ListItem newItem = calendarList.AddItem(newItemInfo);
                    newItem["Title"] = userDisplayName + " is out of Facility";
                    newItem["EventDate"] = DateTime.Now.Date;
                    newItem["EndDate"] = DateTime.Now.Date;
                    newItem["Category"] = "Out Of Facility";
                    newItem["fAllDayEvent"] = true;
                    newItem.Update();
                }
                
                context.ExecuteQuery();

            }

        }
    }
}
