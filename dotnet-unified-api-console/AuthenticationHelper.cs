using System;
using System.Threading.Tasks;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using microsoft.graph;

namespace MicrosoftGraphSampleConsole
{
    public class AuthenticationHelper
    {
        public static string TokenForUser;

        /// <summary>
        /// Async task to acquire token for Application.
        /// </summary>
        /// <returns>Async Token for application.</returns>
        public static async Task<string> AcquireTokenAsyncForApplication()
        {
            return GetTokenForApplication();
        }

        /// <summary>
        /// Get Token for Application.
        /// </summary>
        /// <returns>Token for application.</returns>
        public static string GetTokenForApplication()
        {
            string authority = Constants.AuthString + Constants.TenantId;
            AuthenticationContext authenticationContext = new AuthenticationContext(authority, false);
            // Config for OAuth client credentials 
            ClientCredential clientCred = new ClientCredential(Constants.ClientIdForAppAuthn, Constants.ClientSecret);
            AuthenticationResult authenticationResult = authenticationContext.AcquireToken(Constants.ResourceUrl,
                clientCred);
            string token = authenticationResult.AccessToken;
            return token;
        }

        /// <summary>
        /// Get Active Directory Client for Application.
        /// </summary>
        /// <returns>ActiveDirectoryClient for Application.</returns>
        public static GraphService GetActiveDirectoryClientAsApplication()
        {
            Uri servicePointUri = new Uri(Constants.Url);
            Uri serviceRoot = new Uri(servicePointUri, Constants.TenantId);
            GraphService activeDirectoryClient = new GraphService(serviceRoot,
                async () => await AcquireTokenAsyncForApplication());
            return activeDirectoryClient;
        }

        /// <summary>
        /// Async task to acquire token for User.
        /// </summary>
        /// <returns>Token for user.</returns>
        public static async Task<string> AcquireTokenAsyncForUser()
        {
            return GetTokenForUser();
        }

        /// <summary>
        /// Get Token for User.
        /// </summary>
        /// <returns>Token for user.</returns>
        public static string GetTokenForUser()
        {
            if (TokenForUser == null)
            {
                var redirectUri = new Uri(Constants.redirectUriForUserAuthn);
                string authority = Constants.AuthString + "common";
                AuthenticationContext authenticationContext = new AuthenticationContext(authority, false);
                AuthenticationResult userAuthnResult = authenticationContext.AcquireToken(Constants.ResourceUrl,
                    Constants.ClientIdForUserAuthn, redirectUri, PromptBehavior.RefreshSession);
                TokenForUser = userAuthnResult.AccessToken;
                Console.WriteLine("\n Welcome " + userAuthnResult.UserInfo.GivenName + " " +
                                  userAuthnResult.UserInfo.FamilyName);
            }

            return TokenForUser;
        }

        /// <summary>
        /// Get Active Directory Client for User.
        /// </summary>
        /// <returns>ActiveDirectoryClient for User.</returns>
        public static GraphService GetActiveDirectoryClientAsUser()
        {
            
            Uri serviceRoot = new Uri(Constants.Url);
            GraphService graphClient = new GraphService(serviceRoot, 
                async () => await AcquireTokenAsyncForUser()); 
            return graphClient;
        }
    }
}
