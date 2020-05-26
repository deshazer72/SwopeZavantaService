using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices.ActiveDirectory;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System.Configuration;
using Comprose.Zavanta.PublicApi;
using Comprose.Zavanta.PublicApi.Types;
using System.IO;
using System.Net.Mail;
using System.Net;

namespace SwopeZavantaService
{
    public class Program
    {
        private static String strAddress = ConfigurationManager.AppSettings["serverAddress"];
        private static String strUserName = ConfigurationManager.AppSettings["userName"];
        private static String strPass = ConfigurationManager.AppSettings["password"];
        private static String strTermOUs = ConfigurationManager.AppSettings["termOUs"];
        private static String strZavantaOnlineHost = ConfigurationManager.AppSettings["zavantaHost"];
        private static String strCurrentAccessToken = "";
        private static String strCurrentRefreshToken = "";
        private static String tokenFile = "ZavantaTokens.txt";
        private static Int32 intMinusDays = Convert.ToInt32(ConfigurationManager.AppSettings["intMinusDays"]);
        private static List<string> listEmailNotifications = new List<string>();

        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Zavanta user synchronization starting...");
                GetTokens();
                try
                {
                    SyncAccounts_New();
                    try
                    {
                        SyncAccounts_Updated();
                        try
                        {
                            SyncAccounts_Termed();
                        }
                        catch (Exception e)
                        {
                            SendEmail("Zavanta user sync failure notification", "The program failed while deleting accounts (error: " + e.Message + ")", true);
                        }
                    }
                    catch (Exception e)
                    {
                        SendEmail("Zavanta user sync failure notification", "The program failed while updating accounts (error: " + e.Message + ")", true);
                    }
                }
                catch (Exception e)
                {
                    SendEmail("Zavanta user sync failure notification", "The program failed while creating accounts (error: " + e.Message + ")", true);
                }
            }
            catch (Exception e)
            {
                SendEmail("Zavanta user sync failure notification", "The program failed while acquiring API token (error: " + e.Message + ")", true);
            }

            GenerateFinalIssuesEmail();
            Console.WriteLine("Zavanta user synchronization finished.");

            if (Convert.ToBoolean(ConfigurationManager.AppSettings["pauseAtCompletion"])) //Want this set to false for scheduled process to prevent 
                Console.ReadLine();                                                       //console window from hanging around
        }

        private static void GetTokens()
        {
            strCurrentAccessToken = File.ReadLines(tokenFile).First();
            strCurrentRefreshToken = File.ReadLines(tokenFile).ElementAt(1);
        }

        private static void SyncAccounts_New()
        {
            Account account = new Account();
            var zavantaManagementService = new UserAndGroupManagment(strZavantaOnlineHost, strCurrentAccessToken, strCurrentRefreshToken);
            zavantaManagementService.TokenAcquired += ZavantaManagementService_TokenAcquired;
            DateTime today = DateTime.UtcNow.Date.AddDays(intMinusDays);

            Console.WriteLine("-------------------------------------------");
            Console.WriteLine("Searching active directory for new users...");
            Console.WriteLine("-------------------------------------------");

            using (var context = new PrincipalContext(ContextType.Domain, strAddress, strUserName, strPass))
            {
                using (var searcher = new PrincipalSearcher(new UserPrincipal(context)))
                {
                    foreach (var r in searcher.FindAll())
                    {
                        DirectoryEntry de = r.GetUnderlyingObject() as DirectoryEntry;

                        if (de.Properties["msExchVoiceMailboxID"].Value != null)
                        {

                            int empId = Convert.ToInt32(de.Properties["msExchVoiceMailboxID"].Value);
                            if (empId != 0)
                            {
                                DateTime createdDate = Convert.ToDateTime(de.Properties["whenCreated"].Value);
                                if (createdDate >= today)
                                {
                                    if (de.Properties["givenName"].Value != null)
                                    {
                                        account.FirstName = de.Properties["givenName"].Value.ToString();
                                    }
                                    else
                                    {
                                        account.FirstName = "";
                                    }
                                    if (de.Properties["sn"].Value != null)
                                    {
                                        account.LastName = de.Properties["sn"].Value.ToString();
                                    }
                                    else
                                    {
                                        account.LastName = "";
                                    }
                                    if (de.Properties["department"].Value != null)
                                    {
                                        account.Department = de.Properties["department"].Value.ToString();
                                    }
                                    else
                                    {
                                        account.Department = "";
                                    }
                                    if (de.Properties["mail"].Value != null)
                                    {
                                        account.Email = de.Properties["mail"].Value.ToString();
                                    }
                                    else
                                    {
                                        account.Email = "";
                                    }
                                    if (de.Properties["title"].Value != null)
                                    {
                                        account.Position = de.Properties["title"].Value.ToString();
                                    }
                                    else
                                    {
                                        account.Position = "";
                                    }
                                    if (de.Properties["telephoneNumber"].Value != null)
                                    {
                                        account.Phone = de.Properties["telephoneNumber"].Value.ToString();
                                    }
                                    else
                                    {
                                        account.Phone = "";
                                    }
                                    try
                                    {
                                        zavantaManagementService.CreateUser(email: account.Email, firstName: account.FirstName, lastName: account.LastName, userType: Comprose.Zavanta.PublicApi.UserType.Reader, position: account.Position, department: account.Department, phone: account.Phone);
                                        Console.WriteLine("  User " + account.Email + " created");
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine("   Create of " + account.Email + " user record failed (" + ex.Message + ")");
                                        listEmailNotifications.Add("Create of " + account.Email + " user record failed (" + ex.Message + ")");
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private static void SyncAccounts_Updated()
        {
            Account account = new Account();
            var zavantaManagementService = new UserAndGroupManagment(strZavantaOnlineHost, strCurrentAccessToken, strCurrentRefreshToken);
            zavantaManagementService.TokenAcquired += ZavantaManagementService_TokenAcquired;
            DateTime today = DateTime.UtcNow.Date.AddDays(intMinusDays);
            int Count = 0;

            Console.WriteLine("-----------------------------------------------");
            Console.WriteLine("Searching active directory for updated users...");
            Console.WriteLine("-----------------------------------------------");

            using (var context = new PrincipalContext(ContextType.Domain, strAddress, strUserName, strPass))
            {
                using (var searcher = new PrincipalSearcher(new UserPrincipal(context)))
                {
                    foreach (var r in searcher.FindAll())
                    {
                        DirectoryEntry de = r.GetUnderlyingObject() as DirectoryEntry;

                        if (de.Properties["msExchVoiceMailboxID"].Value != null)
                        {

                            int empId = Convert.ToInt32(de.Properties["msExchVoiceMailboxID"].Value);
                            if (empId != 0)
                            {
                                DateTime updatedDate = Convert.ToDateTime(de.Properties["whenChanged"].Value);

                                if (updatedDate >= today)
                                {
                                    Count++;
                                    if (de.Properties["givenName"].Value != null)
                                    {
                                        account.FirstName = de.Properties["givenName"].Value.ToString();
                                    }
                                    else
                                    {
                                        account.FirstName = "";
                                    }
                                    if (de.Properties["sn"].Value != null)
                                    {
                                        account.LastName = de.Properties["sn"].Value.ToString();
                                    }
                                    else
                                    {
                                        account.LastName = "";
                                    }
                                    if (de.Properties["department"].Value != null)
                                    {
                                        account.Department = de.Properties["department"].Value.ToString();
                                    }
                                    else
                                    {
                                        account.Department = "";
                                    }
                                    if (de.Properties["mail"].Value != null)
                                    {
                                        account.Email = de.Properties["mail"].Value.ToString();
                                    }
                                    else
                                    {
                                        account.Email = "";
                                    }
                                    if (de.Properties["title"].Value != null)
                                    {
                                        account.Position = de.Properties["title"].Value.ToString();
                                    }
                                    else
                                    {
                                        account.Position = "";
                                    }
                                    if (de.Properties["telephoneNumber"].Value != null)
                                    {
                                        account.Phone = de.Properties["telephoneNumber"].Value.ToString();
                                    }
                                    else
                                    {
                                        account.Phone = "";
                                    }

                                    try
                                    {
                                        zavantaManagementService.UpdateUser(email: account.Email, firstName: account.FirstName, lastName: account.LastName, position: account.Position, phone: account.Phone, department: account.Department);
                                        Console.WriteLine("   User " + account.Email + " updated");
                                    }
                                    catch (Exception ex)
                                    {
                                        if (account.Email != "")
                                        {
                                            zavantaManagementService.CreateUser(email: account.Email, firstName: account.FirstName, lastName: account.LastName, userType: Comprose.Zavanta.PublicApi.UserType.Reader, position: account.Position, department: account.Department, phone: account.Phone);
                                            Console.WriteLine("   User " + account.Email + " Created");
                                        }
                                        else
                                        {
                                            Console.WriteLine("   Update of " + account.Email + " user record failed (" + ex.Message + ")");
                                            listEmailNotifications.Add("Update of " + account.Email + " user record failed (" + ex.Message + ")");
                                        }
                                    }

                                    if (Count.Equals(100))
                                    {
                                        System.Threading.Thread.Sleep(60000);  //Zavanta only allows 100 API calls per minute, wait...
                                        Count = 0;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private static void SyncAccounts_Termed()
        {
            Account account = new Account();
            var zavantaManagementService = new UserAndGroupManagment(strZavantaOnlineHost, strCurrentAccessToken, strCurrentRefreshToken);
            zavantaManagementService.TokenAcquired += ZavantaManagementService_TokenAcquired;

            Console.WriteLine("----------------------------------------------");
            Console.WriteLine("Searching active directory for termed users...");
            Console.WriteLine("----------------------------------------------");

            using (var context = new PrincipalContext(ContextType.Domain, strAddress, strUserName, strPass))
            {
                using (var searcher = new PrincipalSearcher(new UserPrincipal(context)))
                {
                    foreach (var r in searcher.FindAll())
                    {
                        DirectoryEntry de = r.GetUnderlyingObject() as DirectoryEntry;
                        if (de.Path.Contains(strTermOUs))
                        {
                            if (de.Properties["mail"].Value != null)
                            {
                                account.Email = de.Properties["mail"].Value.ToString();
                            }
                            else
                            {
                                account.Email = "";
                            }

                            if (account.Email == "")
                            {
                                string firstName = de.Properties["givenName"].Value.ToString();
                                string lastName = de.Properties["sn"].Value.ToString();
                                listEmailNotifications.Add("Unable to delete " + firstName + " " + lastName + " because user record has no email address");
                            }
                            try
                            {
                                zavantaManagementService.DeleteUser(email: account.Email);
                                Console.WriteLine("   User " + account.Email + " deleted");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("   Delete of " + account.Email + " user record failed (" + ex.Message + ")");
                                listEmailNotifications.Add("Delete of " + account.Email + " user record failed (" + ex.Message + ")");
                            }
                        }
                    }
                }
            }
        }

        private static void ZavantaManagementService_TokenAcquired(object sender, TokenAcquiredEventArgs eventArgs)
        {
            using (StreamWriter outputfile = new StreamWriter(Program.tokenFile, false))
            {
                outputfile.WriteLine(eventArgs.Token);
                outputfile.WriteLine(eventArgs.RefreshToken);
            }
            GetTokens();
        }

        private static void GenerateFinalIssuesEmail()
        {
            if (listEmailNotifications.Count > 0)
            {
                string emailBody = "The following issues have occurred in the Zavanta user synchronization process: <br/>";
                foreach (string entry in listEmailNotifications)
                    emailBody += entry + "<br/>";
                SendEmail("Zavanta user sync issue notification", emailBody, false);
            }
        }

        private static void SendEmail(string subject, string bodyText, bool writeBodyToConsole)
        {
            if (writeBodyToConsole)
                Console.WriteLine(bodyText);
            string emailusername = ConfigurationManager.AppSettings["emailFromAddress"].ToString();
            string emailpass = ConfigurationManager.AppSettings["emailFromPassword"].ToString();
            MailMessage mail = new MailMessage();
            mail.From = new System.Net.Mail.MailAddress(ConfigurationManager.AppSettings["emailFromAddress"].ToString());
            SmtpClient smtp = new SmtpClient(ConfigurationManager.AppSettings["emailHost"].ToString(), Convert.ToInt32(ConfigurationManager.AppSettings["emailPort"]));
            smtp.EnableSsl = true;
            smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtp.UseDefaultCredentials = false;
            //smtp.Credentials = new System.Net.NetworkCredential(emailusername, emailpass);
            mail.To.Add(new MailAddress(ConfigurationManager.AppSettings["emailToAddress"].ToString()));
            mail.IsBodyHtml = true;
            mail.Subject = subject;
            mail.Body = bodyText;
            try
            {
                smtp.Send(mail);
            }
            catch (Exception e)
            {
            }
        }
    }
}
