/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using MicrosoftGraphAspNetCoreConnectSample.Helpers;
using Newtonsoft.Json;
using Smartsheet.NET.Core.Entities;
using Smartsheet.NET.Core.Http;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Claims;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace MicrosoftGraphAspNetCoreConnectSample.Controllers
{
    public class HomeController : Controller
    {
        private readonly IConfiguration _configuration;
        private readonly IHostingEnvironment _env;
        private readonly IGraphSdkHelper _graphSdkHelper;        

        public HomeController(IConfiguration configuration, IHostingEnvironment hostingEnvironment, IGraphSdkHelper graphSdkHelper)
        {
            _configuration = configuration;
            _env = hostingEnvironment;
            _graphSdkHelper = graphSdkHelper;            
        }

        [AllowAnonymous]
        // Load user's profile.
        public async Task<IActionResult> Index(string email)
        {
            if (User.Identity.IsAuthenticated)
            {
                // Get users's email.
                email = email ?? User.FindFirst("preferred_username")?.Value;
                ViewData["Email"] = email;

                // Initialize the GraphServiceClient.
                var graphClient = _graphSdkHelper.GetAuthenticatedClient((ClaimsIdentity)User.Identity);

                ViewData["Response"] = await GraphService.GetLicenseInfo(graphClient, email, HttpContext);

                // Reset the current user's email address and the status to display when the page reloads.
                TempData["Message"] = "Got license data.";
            }

            return View();
        }

        [Authorize]
        [HttpPost]
        // Send an email message from the current user.
        public async Task<IActionResult> SendEmail(string recipients)
        {
            if (string.IsNullOrEmpty(recipients))
            {
                TempData["Message"] = "Please add a valid email address to the recipients list!";
                return RedirectToAction("Index");
            }

            try
            {
                // Initialize the GraphServiceClient.
                var graphClient = _graphSdkHelper.GetAuthenticatedClient((ClaimsIdentity)User.Identity);

                // Send the email.
                await GraphService.SendEmail(graphClient, _env, recipients, HttpContext);

                // Reset the current user's email address and the status to display when the page reloads.
                TempData["Message"] = "Success! Your mail was sent.";
                return RedirectToAction("Index");
            }
            catch (ServiceException se)
            {
                if (se.Error.Code == "Caller needs to authenticate.") return new EmptyResult();
                return RedirectToAction("Error", "Home", new { message = "Error: " + se.Error.Message });
            }
        }

        [AllowAnonymous]
        public IActionResult About()
        {
            return View();
        }

        [AllowAnonymous]
        public IActionResult Contact()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [AllowAnonymous]
        public IActionResult Error()
        {
            return View();
        }

        [Authorize]
        public async Task<IActionResult> Loader(string code, string expires_in, string state = "")
        {
            // https://localhost:44334/home/loader?code=y0fpnavbo97ftcf7&expires_in=599952&state=
            // get query string info and get consent code & get auth token from smartsheets            

            string accessToken = HttpContext.Session.GetString("SSAccessToken");
            if (string.IsNullOrWhiteSpace(accessToken))
            {
                accessToken = await ObtainAccessToken("https://api.smartsheet.com/2.0/token", code);
            }

            // call user info.
            SmartsheetHttpClient client = new SmartsheetHttpClient(accessToken, null);
            Smartsheet.NET.Core.Entities.User userInfo = await client.GetCurrentUser(accessToken);

            var details = await client.ListSheets(accessToken);
            StringBuilder sheetDetails = new StringBuilder();;
            foreach(Sheet sheetInfo in details)
            {
                //get published status of the sheet. 
                string s1 = await GetPublishUrl(accessToken, sheetInfo.Id.ToString());
                string s = $"Sheet Name - {sheetInfo.Name}[{sheetInfo.Id}] with RWUrl ={s1} <br>";

                // get list of users the sheet is shared with
                await GetSheetShareInfo(accessToken, sheetInfo.Id.ToString());
                sheetDetails.Append(s);                
            }

            ViewData["Response"] = sheetDetails.ToString();

            // now all that handshake is complete redirect to index page. 
            return View();
        }

        private async Task GetSheetShareInfo(string accessToken, string sheetId)
        {
            string reponseUrl = string.Empty;
            string url = string.Empty;
            if (!string.IsNullOrWhiteSpace(sheetId))
            {
                url = $"https://api.smartsheet.com/2.0/sheets/{sheetId}/shares";
            }

            if (string.IsNullOrWhiteSpace(accessToken))
            {
                throw new Exception("Provided Smartsheet Code cannot be null");
            }

            HttpClient httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Add("Authorization", "Bearer " + accessToken);
            var response = await httpClient.GetAsync(url);

            if (response.IsSuccessStatusCode)
            {
                reponseUrl = await response.Content.ReadAsStringAsync();
                SmartsheetShare shareInfo = JsonConvert.DeserializeObject<SmartsheetShare>(reponseUrl);

                Console.WriteLine("Shared with:" + shareInfo.Data.Count);
            }
        }

        private async Task<string> GetPublishUrl(string accessToken, string sheetId)
        {
            string reponseUrl = string.Empty;
            string url = string.Empty;
            if (!string.IsNullOrWhiteSpace(sheetId))
            {
                url = $"https://api.smartsheet.com/2.0/sheets/{sheetId}/publish ";
            }

            if (string.IsNullOrWhiteSpace(accessToken))
            {
                throw new Exception("Provided Smartsheet Code cannot be null");
            }

            HttpClient httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Add("Authorization", "Bearer " + accessToken);
            var response = await httpClient.GetAsync(url);

            if (response.IsSuccessStatusCode)
            {
                reponseUrl = await response.Content.ReadAsStringAsync();
                Dictionary<string, string> responseInfo = JsonConvert.DeserializeObject<Dictionary<string, string>>(reponseUrl);
                if (responseInfo["readWriteEnabled"].Equals("true", StringComparison.InvariantCultureIgnoreCase))
                {
                    reponseUrl = responseInfo["readWriteUrl"];
                }
                else
                {
                    reponseUrl = string.Empty;
                }
            }

            return reponseUrl;
        }

        private async Task<string> ObtainAccessToken(string url, string code, string clientId= "j6ex9cdw0ci3j9ucs2e", string clientSecret= "g8xirmd07iqr719i2f3", string redirectUri = "")
        {
            if (string.IsNullOrWhiteSpace(code))
            {
                throw new Exception("Provided Smartsheet Code cannot be null");
            }

            var hash = GenerateSHA256String(clientSecret + "|" + code);

            var content = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string, string>("grant_type", "authorization_code"),
                new KeyValuePair<string, string>("client_id", clientId),
                new KeyValuePair<string, string>("code", code),
                new KeyValuePair<string, string>("hash", hash)
            });

            content.Headers.ContentType = new MediaTypeHeaderValue("application/x-www-form-urlencoded");

            HttpClient smartsheetApiClient = new HttpClient();
            var response = await smartsheetApiClient.PostAsync(url, content);
            string authResponse = string.Empty;
            string accessToken = string.Empty;

            if (response.IsSuccessStatusCode)
            {
                authResponse = await response.Content.ReadAsStringAsync();
                Dictionary<string, string> responseInfo = JsonConvert.DeserializeObject<Dictionary<string, string>>(authResponse);
                accessToken = responseInfo["access_token"];          
                HttpContext.Session.SetString("SSAccessToken", accessToken);
            }

            return accessToken;
        }

        private static string GenerateSHA256String(string inputString)
        {
            SHA256 sha256 = SHA256.Create();
            byte[] bytes = Encoding.UTF8.GetBytes(inputString);
            byte[] hash = sha256.ComputeHash(bytes);

            StringBuilder result = new StringBuilder();
            for (int i = 0; i < hash.Length; i++) { result.Append(hash[i].ToString("X2")); }

            return result.ToString();            
        }

    }
}
