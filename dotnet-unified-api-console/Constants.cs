namespace MicrosoftGraphSampleConsole
{
    internal class Constants
    {
        // To run this console app you need to have 2 application registrations - as the app runs in 2 modes...

        // app only configuration
        // In the Azure management portal, register a web application, and get a clientId and secret/key, 
        // and configure permissions according to README
        public const string TenantId = "[Enter your TenantId.  Looks a bit like:  c437555e-20fc-4d8e-97cc-70f7d270a8ba]";
        public const string ClientIdForAppAuthn = "[Enter your client Id.  Looks a bit like:  32c0408e-2a40-4e75-88b6-c0a3497de749 ]";
        public const string ClientSecret = "Enter your Secret/Key.  Looks a bit like: DxD0Rv96rEvo0QUBy7qF7nmJdsV/XcJaNoc3Sf/IHBQ= ]";

        // app+user console app configuration
        // In the Azure management portal, register a native client application, 
        // and configure permissions according to README
        public const string ClientIdForUserAuthn = "[Enter your client Id.  Looks a bit like:  b237040e-20a0-48b6-85e3-9186b7c1eef7 ]";
        public const string redirectUriForUserAuthn = "[Enter your redirect Uri for the app.  Looks a bit like: http://localhost:44323 ]";

        public const string AuthString = "https://login.microsoftonline.com/";
        public const string ResourceUrl = "https://graph.microsoft.com/";
        public const string Url = "https://graph.microsoft.com/v1.0/";

    }
}