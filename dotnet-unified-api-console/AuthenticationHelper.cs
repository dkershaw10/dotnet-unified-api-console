using System;
using System.Threading.Tasks;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Graph;

namespace MicrosoftGraphSampleConsole
{
    public class AuthenticationHelper
    {
        public static string TokenForUser;

        /// <summary>
        /// Get Active Directory Client for Application.
        /// </summary>
        /// <returns>ActiveDirectoryClient for Application.</returns>
        public static GraphServiceClient GetActiveDirectoryClientAsApplication()
        {
            Uri servicePointUri = new Uri(Constants.Url);
            Uri serviceRoot = new Uri(servicePointUri, Constants.TenantId);
            return new GraphServiceClient(
                new AppConfig
                {
                    ActiveDirectoryAppId = Constants.ClientIdForAppAuthn,
                    ActiveDirectoryClientSecret = Constants.ClientSecret,
                    ActiveDirectoryReturnUrl = Constants.redirectUriForUserAuthn,
                    ActiveDirectoryServiceEndpointUrl = Constants.Url,
                    ActiveDirectoryServiceResource = Constants.ResourceUrl,
                },
                new AdalCredentialCache(),
                new HttpProvider(),
                new AdalServiceInfoProvider(),
                ClientType.Business);
        }

        /// <summary>
        /// Async task to acquire token for User.
        /// </summary>
        /// <returns>Token for user.</returns>
        public static Task<string> AcquireTokenAsyncForUser()
        {
            return Task.FromResult(GetTokenForUser());
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
        public static GraphServiceClient GetActiveDirectoryClientAsUser()
        {
            return new GraphServiceClient(
                new AppConfig
                {
                    ActiveDirectoryAppId = Constants.ClientIdForUserAuthn,
                    //ActiveDirectoryClientSecret = clientSecret,
                    ActiveDirectoryReturnUrl = Constants.redirectUriForUserAuthn,
                    ActiveDirectoryServiceEndpointUrl = Constants.Url,
                    ActiveDirectoryServiceResource = Constants.ResourceUrl,
                },
                new AdalCredentialCache(),
                new HttpProvider(),
                new AdalServiceInfoProvider(),
                ClientType.Business);
        }
    }
}
