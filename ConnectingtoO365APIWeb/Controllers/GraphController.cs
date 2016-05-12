using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web.Http;

namespace ConnectingtoO365APIWeb.Controllers
{
    public class GraphController : ApiController
    {
        [HttpGet]
        [Route("api/Graph/GetMeProfile")]
        public async Task<HttpResponseMessage> GetMeProfile(string connectionId)
        {
            //
            // Retrieve the contact based on the contactId, in this case their email.
            //

            // Get the user's access token
            OAuthController.UserAccessDetails accessDetails = OAuthController.GetUsersAccessDetails(connectionId);

            HttpClient client = new HttpClient();

            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, @"https://graph.microsoft.com/v1.0/me/");

            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessDetails.AccessToken);

            HttpResponseMessage response = await client.SendAsync(request);

            return response;
        }

        [HttpGet]
        [Route("api/Graph/GetPeople")]
        public async Task<HttpResponseMessage> GetPeople(string connectionId)
        {
            //
            // Retrieve the contact based on the contactId, in this case their email.
            //

            // Get the user's access token
            OAuthController.UserAccessDetails accessDetails = OAuthController.GetUsersAccessDetails(connectionId);

            HttpClient client = new HttpClient();

            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, @"https://graph.microsoft.com/beta/me/people");

            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessDetails.AccessToken);

            HttpResponseMessage response = await client.SendAsync(request);

            return response;
        }

    }
}
