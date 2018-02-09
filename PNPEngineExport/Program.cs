using Microsoft.SharePoint.Client;
using PNPEngineFunctions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PNPEngineExport.Templates;

namespace PNPEngineExport
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                LogWriter.Current.WriteLine("Beginning Site Configuration");

                ConnectionSettings connectionSettings = AskForConnectionSettings();
                Console.Write("Path du xml ( ex: C:\\MyTemplates ) : ");
                string pathXml = Console.ReadLine();
                Console.Write("Fichier xml ( ex: template.xml ) : ");
                string nameXml = Console.ReadLine();

                if (connectionSettings != null)
                {
                    LogWriter.Current.WriteLine("Connecting to site collection");
                    using (ClientContext context = new O365LoginPasswdSiteCollectionConnection(connectionSettings).Connect())
                    {
                        Web rootWeb = context.Site.RootWeb;
                        context.Load(rootWeb);
                        context.Load(rootWeb, w => w.AllProperties);
                        context.ExecuteQueryRetry();
                        LogWriter.Current.WriteLine("Extraction du site template");
                        TemplatesExtractor.Apply(rootWeb, pathXml, nameXml);
                    }
                    LogWriter.Current.WriteLine("Extraction Terminé");
                    Console.ReadLine();
                }
            }
            catch (Exception ex)
            {
                LogWriter.Current.WriteLine(ex.Message);
                LogWriter.Current.WriteLine(ex.StackTrace);
                if (ex.InnerException != null)
                {
                    LogWriter.Current.WriteLine(ex.InnerException.Message);
                    LogWriter.Current.WriteLine(ex.InnerException.StackTrace);
                }
                Console.ReadLine();
            }
            finally
            {
                LogWriter.Current.UpdateLogFile("PNPEngineExport.log");
            }
        }

        private static ConnectionSettings AskForConnectionSettings()
        {

            Console.Write("URL absolue du site : ");
            string siteUrl = Console.ReadLine();

            Console.Write("Login : ");
            string login = Console.ReadLine();

            Console.Write("Mot de passe : ");
            string password = string.Empty;
            while (true)
            {
                ConsoleKeyInfo i = Console.ReadKey(true);
                if (i.Key == ConsoleKey.Enter)
                {
                    break;
                }
                else if (i.Key == ConsoleKey.Backspace)
                {
                    if (password.Length > 0)
                    {
                        password.Remove(password.Length - 1);
                        Console.Write("\b \b");
                    }
                }
                else
                {
                    password += i.KeyChar;
                    Console.Write("*");
                }
            }
            LogWriter.Current.WriteLine("");
            Console.Write("useAppAuthentication : ");
            Boolean useAppAuthentication = false;
            string useAppAuthenticationString = Console.ReadLine();
            if (useAppAuthenticationString.ToLower() == "yes" || useAppAuthenticationString.ToLower() == "true")
            {
                useAppAuthentication = true;
            }


            return new ConnectionSettings
            {
                Login = login,
                Password = password,
                SiteUrl = siteUrl,
                UseAppAuthentication = useAppAuthentication
            };
        }
    }
}
