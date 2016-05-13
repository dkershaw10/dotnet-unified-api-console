namespace MicrosoftGraphSampleConsole
{
    internal class Constants
    {
        // To run this console app you need to have 2 application registrations - as the app runs in 2 modes...

        // app only configuration
        // In the Azure management portal, register a web application, and get a clientId and secret/key, 
        // and configure permissions according to README
        public const string TenantId = "7b0208e6-0f64-47ae-bb8b-3e1ba480e347";
        public const string ClientIdForAppAuthn = "336057fa-0c14-46ae-a59b-32d81c531c67";
        public const string ClientSecret = "DxD0Rv92rEvo0QUBy7qF3nmJdsV/XcJaNoc3Sf/IHBQ=";

        // app+user console app configuration
        // In the Azure management portal, register a native client application, 
        // and configure permissions according to README
        public const string ClientIdForUserAuthn = "0b45d997-3470-40ec-ad2c-a9b816e076df";
        public const string redirectUriForUserAuthn = "http://localhost:44323";

        public const string AuthString = "https://login.microsoftonline.com/";
        public const string ResourceUrl = "https://graph.microsoft.com/";
        public const string Url = "https://graph.microsoft.com/v1.0/";

    }
}