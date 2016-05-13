#region

using System;
using System.Net;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Threading.Tasks;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Net.Http;

#endregion

namespace MicrosoftGraphSampleConsole
{
    public class Program
    {

        // Single-Threaded Apartment required for OAuth2 Authz Code flow (User Authn) to execute for this demo app
        [STAThread]
        private static void Main()
        {
            // record start DateTime of execution
            string currentDateTime = DateTime.Now.ToUniversalTime().ToString();

            ServicePointManager.ServerCertificateValidationCallback += (sender, cert, chain, sslPolicyErrors) => true;

            Console.WriteLine("Run operations for signed-in user, or in app-only mode.\n");
            Console.WriteLine("[a] - app-only\n[u] - as user\n[b] - both as user first, and then as app.\nPlease enter your choice:\n");

            ConsoleKeyInfo key = Console.ReadKey();
            switch (key.KeyChar)
            {
                case 'a':
                    Console.WriteLine("\nRunning app-only mode\n\n");
                    Requests.AppMode();
                    break;
                case 'b':
                    Console.WriteLine("\nRunning app-only mode, followed by user mode\n\n");
                    Requests.AppMode();
                    var appModeTask = Task.Run(() => Requests.UserMode());
                    appModeTask.Wait();
                    break;
                case 'u':
                    Console.WriteLine("\nRunning in user mode\n\n");
                    var userModeTask = Task.Run(() => Requests.UserMode());
                    userModeTask.Wait();
                    break;
                default:
                    Console.WriteLine("\nSelection not recognized. Running in user mode\n\n");
                    break;
            }

            //*********************************************************************************************
            // End of Demo Console App
            //*********************************************************************************************

            Console.WriteLine("\nCompleted at {0} \n Press Any Key to Exit.", currentDateTime);
            Console.ReadKey();
        }
    }
}
