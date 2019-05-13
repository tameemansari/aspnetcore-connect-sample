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
using Smartsheet.NET.Core.Entities;
using Smartsheet.NET.Core.Http;
using System.Security.Claims;
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
                accessToken = await SmartSheetHelper.ObtainAccessToken("https://api.smartsheet.com/2.0/token", code);
                HttpContext.Session.SetString("SSAccessToken", accessToken);
            }

            // call user info.
            SmartsheetHttpClient client = new SmartsheetHttpClient(accessToken, null);
            Smartsheet.NET.Core.Entities.User userInfo = await client.GetCurrentUser(accessToken);

            var details = await client.ListSheets(accessToken);
            StringBuilder sheetDetails = new StringBuilder();;
            foreach(Sheet sheetInfo in details)
            {
                //get published status of the sheet. 
                string s1 = await SmartSheetHelper.GetPublishUrl(accessToken, sheetInfo.Id.ToString());
                string s = $"Sheet Name - {sheetInfo.Name}[{sheetInfo.Id}] with RWUrl ={s1} <br>";

                // get list of users the sheet is shared with
                await SmartSheetHelper.GetSheetShareInfo(accessToken, sheetInfo.Id.ToString());
                sheetDetails.Append(s);                
            }

            ViewData["Response"] = sheetDetails.ToString();

            // now all that handshake is complete redirect to index page. 
            return View();
        }


    }
}
