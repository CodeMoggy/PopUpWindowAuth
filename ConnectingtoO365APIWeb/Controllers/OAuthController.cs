// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
using System;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Web.Http;

using Newtonsoft.Json;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using ConnectingtoO365APIWeb.Cache;
using Microsoft.AspNet.SignalR;
using ConnectingtoO365APIWeb.Hubs;
using Microsoft.AspNet.SignalR.Hubs;
using System.Configuration;

namespace ConnectingtoO365APIWeb.Controllers
{
    public class OAuthController : OAuthControllerWithHub<OAuthHub>
    {
        // Register the app in Azure AD to get these values.
        // Client ID is found on the "Configure" tab for the application in Azure Management Portal
        private static string ClientId = ConfigurationManager.AppSettings["ida:ClientId"]; 

        // Client Secret is generated on the "Configure" tab for the application in Azure Management Portal,
        // Under "Keys"
        private static string ClientSecret = ConfigurationManager.AppSettings["ida:ClientSecret"];

        // ERPContactService unique identifier 
        private static string Resource = ConfigurationManager.AppSettings["ida:ResourceId"];

        // OAuth endpoints
        private const string OAuthUrl = "https://login.microsoftonline.com/{0}";

        // multi-tenant - replace common with tenant name for single tenant
        private static readonly string AuthorizeUrlNoResource = string.Format(CultureInfo.InvariantCulture,
            OAuthUrl,
            "common/oauth2/authorize?response_type=code&client_id={0}&redirect_uri={1}&state={2}");
        private static readonly Uri RedirectUrl = new Uri(
            System.Web.HttpContext.Current.Request.Url, "/Addin/oauthRedirect.html");


        [HttpGet()]
        [Route("api/OAuth/GetAuthorizationUrl")]
        public string GetAuthorizationUrl(string connectionId)
        {
            // Generate a new GUID to add to the request.
            // Save the GUID mapped to the user, so we can look up the user
            // once we have the auth response.
            string stateGuid = connectionId;
            AppConfigCache.AddStateGuid(stateGuid, connectionId);

            return String.Format(CultureInfo.InvariantCulture,
            AuthorizeUrlNoResource,
            Uri.EscapeDataString(ClientId),
            Uri.EscapeDataString(RedirectUrl.ToString()),
            Uri.EscapeDataString(stateGuid));
        }

        
        [HttpPost()]
        public string CompleteOAuthFlow(AuthorizationParameters parameters)
        {
            // Look up the email from the guid/user map.
            string userId = AppConfigCache.GetUserFromStateGuid(parameters.State);

            if (string.IsNullOrEmpty(userId))
            {
                // Per the Azure docs, the response from the auth code request has
                // to include the value of the state parameter passed in the request.
                // If it is not the same, then you should not accept the response.
                throw new HttpResponseException(Request.CreateErrorResponse(HttpStatusCode.OK,
                    "Unknown state returned in OAuth flow."));
            }

            try
            {
                ClientCredential credential = new ClientCredential(ClientId, ClientSecret);
                string authority = string.Format(CultureInfo.InvariantCulture, OAuthUrl, "common");
                AuthenticationContext authContext = new AuthenticationContext(authority);
                AuthenticationResult result = authContext.AcquireTokenByAuthorizationCode(
                    parameters.Code, new Uri(RedirectUrl.GetLeftPart(UriPartial.Path)), credential, Resource);

                // Cache the refresh token
                var appConfig = new AppConfig();
                appConfig.RefreshToken = result.RefreshToken;

                // Save the user's configuration in our confic cache
                AppConfigCache.AddUserConfig(parameters.State, appConfig);

                // notify client side that the oauth flow is complete and therefore the contact information can be retrieved
                Hub.Clients.Group(parameters.State).completed();

                return "OAuth succeeded. Please close this window to continue.";
            }
            catch (AdalServiceException ex)
            {
                throw new HttpResponseException(Request.CreateErrorResponse(HttpStatusCode.OK,
                    "OAuth failed. " + ex.ToString()));
            }
        }

        public static UserAccessDetails GetUsersAccessDetails(string connectionId)
        {
            try
            {
                // Get the user's config, which contains the access token
                AppConfig appConfig = AppConfigCache.GetUserConfig(connectionId);

                // Request authorization for OneDrive
                string authority = string.Format(CultureInfo.InvariantCulture, OAuthUrl, "common");
                AuthenticationContext authContext = new AuthenticationContext(authority);
                ClientCredential credential = new ClientCredential(ClientId, ClientSecret);
                AuthenticationResult result = authContext.AcquireTokenByRefreshToken(appConfig.RefreshToken, credential, Resource);

                // Update refresh token
                appConfig.RefreshToken = result.RefreshToken;
                AppConfigCache.AddUserConfig(connectionId, appConfig);

                return new UserAccessDetails()
                {
                    AccessToken = result.AccessToken
                };
            }
            catch (AdalException)
            {
                return null;
            }
        }

        #region Helper classes

        public class AuthorizationParameters
        {
            public string Code { get; set; }
            public string State { get; set; }
        }

        public class UserAccessDetails
        {
            public string AccessToken { get; set; }
        }
        #endregion
    }
}

// MIT License: 

// Permission is hereby granted, free of charge, to any person obtaining 
// a copy of this software and associated documentation files (the 
// ""Software""), to deal in the Software without restriction, including 
// without limitation the rights to use, copy, modify, merge, publish, 
// distribute, sublicense, and/or sell copies of the Software, and to 
// permit persons to whom the Software is furnished to do so, subject to 
// the following conditions: 

// The above copyright notice and this permission notice shall be 
// included in all copies or substantial portions of the Software. 

// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND, 
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF 
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND 
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE 
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION 
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION 
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE. 