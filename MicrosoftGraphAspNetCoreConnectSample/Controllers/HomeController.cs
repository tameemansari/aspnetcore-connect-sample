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
using MicrosoftGraphAspNetCoreConnectSample.Helpers.Models;
using Smartsheet.NET.Core.Entities;
using Smartsheet.NET.Core.Http;
using System.Collections.Generic;
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
            
            SmartsheetHttpClient client = new SmartsheetHttpClient(accessToken, null);
            var details = await client.ListSheets(accessToken);

            string email = User.FindFirst("preferred_username")?.Value;
            var graphClient = _graphSdkHelper.GetAuthenticatedClient((ClaimsIdentity)User.Identity);
            StringBuilder sheetDetails = new StringBuilder();;
            foreach(Sheet thisSheet in details)
            {
                string sheetPublishUrl = await SmartSheetHelper.GetPublishUrl(accessToken, thisSheet.Id.ToString());

                SheetInformation sheetToBuild = new SheetInformation()
                {
                    SheetId = thisSheet.Id.ToString(),
                    SheetName = thisSheet.Name,
                    SheetRWUrl = sheetPublishUrl,
                    Collaborators = new List<string>(),
                };
                
                // get list of users the sheet is shared with
                SmartSheetShare sheetSharedWith = await SmartSheetHelper.GetSheetShareInfo(accessToken, thisSheet.Id.ToString());
                if (sheetSharedWith != null)
                {
                    List<string> userInfo = new List<string>();
                    foreach(Datum datum in sheetSharedWith.Data)
                    {
                        if (!string.IsNullOrWhiteSpace(datum.Email))
                        {
                            sheetToBuild.Collaborators.Add(datum.Email);
                        }
                    }                    
                }
                await GraphService.CreateGroupAndTeamApp(graphClient, email, sheetToBuild);

                //get published status of the sheet.                 
                string s = $"Sheet Name - {thisSheet.Name}[{thisSheet.Id}] with RWUrl ={sheetPublishUrl} <br>";
                sheetDetails.Append(s);                
            }

            ViewData["Response"] = sheetDetails.ToString();

            // now all that handshake is complete redirect to index page. 
            return View();
        }


    }
}
