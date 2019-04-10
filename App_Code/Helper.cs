using System;
using System.Collections.Generic;
using System.Web;
using System.Globalization;

using System.IO;
using System.Net.Mail;
using System.Web.UI;
//using System.Linq;
using System.Configuration;
using System.Web.Configuration;
using System.Net.Configuration;


namespace Automate.Utilities
{
    public static partial class Helpers
    {
        private static IFormatProvider enCulture = new CultureInfo("en-US");
        private static IFormatProvider thCulture = new CultureInfo("th-TH");
        ////private string reportServer = ConfigurationManager.AppSettings["ReportServer"];
        ////private string reportDatabase = ConfigurationManager.AppSettings["ReportDatabase"];
        ////private string reportUserName = ConfigurationManager.AppSettings["ReportUserName"];
        ////private string reportPassword = ConfigurationManager.AppSettings["ReportPassword"];


        public static IFormatProvider ENCulture
        {
            get { return enCulture; }
        }
        public static IFormatProvider THCulture
        {
            get { return thCulture; }
        }
        public static string ReportServer
        {
            get { return ConfigurationManager.AppSettings["ReportServer"]; }
        }
        public static string ReportDatabase
        {
            get { return ConfigurationManager.AppSettings["ReportDatabase"]; }
        }
        public static string ReportUserName
        {
            get { return ConfigurationManager.AppSettings["ReportUserName"]; }
        }
        public static string ReportPassword
        {
            get { return ConfigurationManager.AppSettings["ReportPassword"]; }
        }

        public static bool SendHTMLMail(MailMessage msg)
        {
            ////Configuration configurationFile = WebConfigurationManager.OpenWebConfiguration(HttpContext.Current.Server.MapPath("~/web.config"));
            ////Configuration configurationFile = WebConfigurationManager.OpenWebConfiguration(HttpContext.Current.Server.MapPath("~/web.config"));
            ////MailSettingsSectionGroup mailSettings = configurationFile.GetSectionGroup("system.net/mailSettings") as MailSettingsSectionGroup;

            ////if (mailSettings != null)
            ////{
            ////    try
            ////    {
            ////        SmtpClient smtp = new SmtpClient();
            ////        smtp.Host = mailSettings.Smtp.Network.Host;
            ////        smtp.Port = mailSettings.Smtp.Network.Port;
            ////        smtp.Credentials = new System.Net.NetworkCredential(mailSettings.Smtp.Network.UserName
            ////                                                           , mailSettings.Smtp.Network.Password);
            ////        smtp.EnableSsl = false;
            ////        smtp.Send(msg);
            ////    }
            ////    catch (Exception ex)
            ////    {
            ////        throw new Exception("Error Infomation ", ex);
            ////    }
            ////}

            bool success = false;
            System.Net.ServicePointManager.ServerCertificateValidationCallback =
                delegate(object s, System.Security.Cryptography.X509Certificates.X509Certificate certificate,
                System.Security.Cryptography.X509Certificates.X509Chain chain, System.Net.Security.SslPolicyErrors sslPolicyErrors)
                { return true; };
            SmtpClient smtp = new SmtpClient();
            smtp.Host = "smtp.gmail.com";
            smtp.UseDefaultCredentials = true;
            smtp.EnableSsl = true;
            smtp.Credentials = new System.Net.NetworkCredential("eakkaphol.s@gmail.com", "eak123**");
            smtp.Port = 587;


            try
            {
                smtp.Send(msg);
                success = true;
            }
            catch
            {
                success = false;
            }
            return success;

        }


    }
}
