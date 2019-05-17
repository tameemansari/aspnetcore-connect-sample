/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.Graph;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace MicrosoftGraphAspNetCoreConnectSample.Helpers
{
    public static class GraphService
    {
        // Load user's profile in formatted JSON.
        public static async Task<string> GetUserJson(GraphServiceClient graphClient, string email, HttpContext httpContext)
        {
            if (email == null) return JsonConvert.SerializeObject(new { Message = "Email address cannot be null." }, Formatting.Indented);

            try
            {
                // Load user profile.
                var user = await graphClient.Users[email].Request().GetAsync();
                return JsonConvert.SerializeObject(user, Formatting.Indented);
            }
            catch (ServiceException e)
            {
                switch (e.Error.Code)
                {
                    case "Request_ResourceNotFound":
                    case "ResourceNotFound":
                    case "ErrorItemNotFound":
                    case "itemNotFound":
                        return JsonConvert.SerializeObject(new { Message = $"User '{email}' was not found." }, Formatting.Indented);
                    case "ErrorInvalidUser":
                        return JsonConvert.SerializeObject(new { Message = $"The requested user '{email}' is invalid." }, Formatting.Indented);
                    case "AuthenticationFailure":
                        return JsonConvert.SerializeObject(new { e.Error.Message }, Formatting.Indented);
                    case "TokenNotFound":
                        await httpContext.ChallengeAsync();
                        return JsonConvert.SerializeObject(new { e.Error.Message }, Formatting.Indented);
                    default:
                        return JsonConvert.SerializeObject(new { Message = "An unknown error has occurred." }, Formatting.Indented);
                }
            }
        }

        public static async Task<string> GetLicenseInfo(GraphServiceClient graphClient, string upn, HttpContext httpContext)
        {
            string licenseInfo = string.Empty;
            try
            {
                var userInfo = await graphClient.Users[upn].Request().GetAsync().ConfigureAwait(false);

                // for the passed in userObjectId assign 
                bool alreadyAssigned = false;
                string m365skuId = "cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46";

                var response1 = await graphClient.Users[userInfo.Id].LicenseDetails.Request().GetAsync().ConfigureAwait(false);
                if (response1.Count >= 0)
                {
                    foreach (LicenseDetails ld in response1)
                    {
                        string tocheck = ld.SkuId.ToString();
                        if (m365skuId.Equals(tocheck, StringComparison.OrdinalIgnoreCase))
                        {
                            alreadyAssigned = true;
                            break;
                        }
                    }
                }

                if (alreadyAssigned) licenseInfo += "License to this user is already assigned.";
                else
                {
                    licenseInfo += "Currently user is not assigned a valid license... attempting to assign license after updating usage location..";
                    var userToUpdate = await graphClient.Users[userInfo.Id].Request().GetAsync();

                    if (string.IsNullOrWhiteSpace(userToUpdate.UsageLocation))
                    {
                        userToUpdate.UsageLocation = "US";
                        await graphClient.Users[userInfo.Id].Request().UpdateAsync(userToUpdate);
                    }
                    licenseInfo += "... updated usage location... assigninging license now...";


                    AssignedLicense aLicense = new AssignedLicense { SkuId = new Guid(m365skuId) };
                    IList<AssignedLicense> licensesToAdd = new AssignedLicense[] { aLicense };
                    IList<Guid> licensesToRemove = Array.Empty<Guid>();

                    await graphClient.Users[userInfo.Id].AssignLicense(licensesToAdd, licensesToRemove).Request().PostAsync().ConfigureAwait(false);
                    licenseInfo += "... Assignment success.";
                }

                // licenseInfo += await CreateGroupAndTeamApp(graphClient, upn);

                return JsonConvert.SerializeObject(licenseInfo, Formatting.Indented);
            }
            catch (ServiceException e)
            {
                return JsonConvert.SerializeObject(new { e.Error.Message }, Formatting.Indented);
            }
        }

        public static async Task<string> CreateGroupAndTeamApp(GraphServiceClient graphClient, SheetInformation sheetInformation)
        {
            try
            {                
                string suffixInfo = Guid.NewGuid().ToString().Substring(0, 8);
                string sheetName = sheetInformation.SheetName;
                
                #region Create group and add executing user as owner and append members. 
                var grpInfo = await graphClient.Groups.Request().AddAsync(new Group()
                {                    
                    DisplayName = $"{sheetName} Sheet Collaborators-{suffixInfo}", 
                    MailNickname = $"grp1{sheetInformation.SheetId}-{suffixInfo}",
                    Description = $"Team for collaborating on {sheetName} smartsheet.",
                    Visibility = "Private",
                    GroupTypes = new List<string>() { "Unified" },
                    MailEnabled = true,
                    SecurityEnabled = false,
                });               
                
                // append members from sheetInformation. 
                foreach(string userUpn in sheetInformation.Collaborators)
                {
                    Debug.WriteLine($"Appending {userUpn}");
                    if (!string.IsNullOrWhiteSpace(userUpn))
                    {
                        var memberToAppend = await graphClient.Users[userUpn].Request().GetAsync().ConfigureAwait(false);
                        try { await graphClient.Groups[grpInfo.Id].Members.References.Request().AddAsync(memberToAppend); } catch { }                        
                    }
                }
                #endregion

                // TODO :: This is an expensive call.
                #region Build team and channel.
                var teamInfo = await graphClient.Groups[grpInfo.Id].Team.Request().PutAsync(new Team()
                {
                    MemberSettings = new TeamMemberSettings() { AllowCreateUpdateChannels = true, },
                    MessagingSettings = new TeamMessagingSettings() { AllowUserEditMessages = true, AllowUserDeleteMessages = true, },
                });

                // build channel into team.
                var channelInfo = await graphClient.Teams[grpInfo.Id].Channels.Request().AddAsync(new Channel()
                {
                    DisplayName = $"{sheetName}-{suffixInfo}",
                    Description = "Proton Channel description",                    
                });
                #endregion

                #region Check and install app to team. 
                var teamApps = await graphClient.AppCatalogs.TeamsApps.Request().GetAsync();
                TeamsApp sheetApp = teamApps.First(x => x.Id.Equals("f4d81e8e-4500-44c2-8328-9e06cbe037c5", StringComparison.InvariantCultureIgnoreCase));

                // :: TODO :: removed for perf optimization. Not required. 
                // var installedApps = await graphClient.Teams[grpInfo.Id].InstalledApps.Request().Expand("teamsAppDefinition").GetAsync();
                // bool isSmartsheetsInstalled = false; // installedApps.Any(x => x.TeamsAppDefinition.TeamsAppId.Equals("f4d81e8e-4500-44c2-8328-9e06cbe037c5", StringComparison.InvariantCultureIgnoreCase));

                bool isSmartsheetsInstalled = false; 
                if (!isSmartsheetsInstalled)
                {
                    TeamsAppInstallation appInstall = new TeamsAppInstallation()
                    {
                        AdditionalData = new Dictionary<string, object>()
                        {
                            {
                                "teamsApp@odata.bind", "https://graph.microsoft.com/beta/appCatalogs/teamsApps/f4d81e8e-4500-44c2-8328-9e06cbe037c5"
                            }
                        },
                    };
                    await graphClient.Teams[grpInfo.Id].InstalledApps.Request().AddAsync(appInstall);
                }
                #endregion

                #region Pin smartsheet app to channel with right configuration.                
                //one didnt bind - https://app.smartsheet.com/b/publish?EQBCT=f7615490df8a44238ddc286745ade920&ss_src=mst                
                var sheetTab = await graphClient.Teams[grpInfo.Id].Channels[channelInfo.Id].Tabs.Request().AddAsync(new TeamsTab()
                {
                    AdditionalData = new Dictionary<string, object>()
                    {
                        {
                            "teamsApp@odata.bind", "https://graph.microsoft.com/beta/appCatalogs/teamsApps/f4d81e8e-4500-44c2-8328-9e06cbe037c5"
                        }
                    },                    
                    DisplayName = sheetName,
                    TeamsApp = sheetApp,
                    Configuration = new TeamsTabConfiguration()
                    {
                        ContentUrl = sheetInformation.SheetRWUrl,
                        WebsiteUrl = sheetInformation.SheetRWUrl,
                    },
                });
                #endregion
                
                // return the url to created team. 
                return sheetTab.WebUrl; 
            }
            catch (ServiceException)
            {
                return string.Empty;
            }
        }

        public static async Task<string> AddUserToTeam(GraphServiceClient graphClient, string groupId, string userAlias, string userDomain)
        {
            string m365skuId = "cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46";
            try
            {
                PasswordProfile passwordProfile = new PasswordProfile
                {
                    Password = "Password!1",
                    ForceChangePasswordNextSignIn = false
                };

                User user = new User
                {
                    DisplayName = $"{userAlias} displayName",
                    UserPrincipalName = $"{userAlias}@{userDomain}",
                    MailNickname = userAlias,
                    AccountEnabled = true,
                    UsageLocation = "US",
                    PasswordProfile = passwordProfile
                };

                // add user. 
                var createdUser = await graphClient.Users.Request().AddAsync(user);

                AssignedLicense aLicense = new AssignedLicense { SkuId = new Guid(m365skuId) };
                IList<AssignedLicense> licensesToAdd = new AssignedLicense[] { aLicense };
                IList<Guid> licensesToRemove = Array.Empty<Guid>();

                // assign license. 
                await graphClient.Users[createdUser.Id].AssignLicense(licensesToAdd, licensesToRemove).Request().PostAsync().ConfigureAwait(false);

                // add user to group.
                await graphClient.Groups[groupId].Members.References.Request().AddAsync(createdUser);

                return JsonConvert.SerializeObject("Success", Formatting.Indented);
            }
            catch (ServiceException se)
            {
                return JsonConvert.SerializeObject(se.Message, Formatting.Indented);
            }

        }
        public static async Task<Stream> GetPictureStream(GraphServiceClient graphClient, string email, HttpContext httpContext)
        {
            if (email == null) throw new Exception("EmailIsNull");

            Stream pictureStream = null;

            try
            {
                try
                {
                    // Load user's profile picture.
                    pictureStream = await graphClient.Users[email].Photo.Content.Request().GetAsync();
                }
                catch (ServiceException e)
                {
                    if (e.Error.Code == "GetUserPhoto") // User is using MSA, we need to use beta endpoint
                    {
                        // Set Microsoft Graph endpoint to beta, to be able to get profile picture for MSAs 
                        graphClient.BaseUrl = "https://graph.microsoft.com/beta";

                        // Get profile picture from Microsoft Graph
                        pictureStream = await graphClient.Users[email].Photo.Content.Request().GetAsync();

                        // Reset Microsoft Graph endpoint to v1.0
                        graphClient.BaseUrl = "https://graph.microsoft.com/v1.0";
                    }
                }
            }
            catch (ServiceException e)
            {
                switch (e.Error.Code)
                {
                    case "Request_ResourceNotFound":
                    case "ResourceNotFound":
                    case "ErrorItemNotFound":
                    case "itemNotFound":
                    case "ErrorInvalidUser":
                        // If picture not found, return the default image.
                        throw new Exception("ResourceNotFound");
                    case "TokenNotFound":
                        await httpContext.ChallengeAsync();
                        return null;
                    default:
                        return null;
                }
            }

            return pictureStream;
        }
        public static async Task<Stream> GetMyPictureStream(GraphServiceClient graphClient, HttpContext httpContext)
        {
            Stream pictureStream = null;

            try
            {
                try
                {
                    // Load user's profile picture.
                    pictureStream = await graphClient.Me.Photo.Content.Request().GetAsync();
                }
                catch (ServiceException e)
                {
                    if (e.Error.Code == "GetUserPhoto") // User is using MSA, we need to use beta endpoint
                    {
                        // Set Microsoft Graph endpoint to beta, to be able to get profile picture for MSAs 
                        graphClient.BaseUrl = "https://graph.microsoft.com/beta";

                        // Get profile picture from Microsoft Graph
                        pictureStream = await graphClient.Me.Photo.Content.Request().GetAsync();

                        // Reset Microsoft Graph endpoint to v1.0
                        graphClient.BaseUrl = "https://graph.microsoft.com/v1.0";
                    }
                }
            }
            catch (ServiceException e)
            {
                switch (e.Error.Code)
                {
                    case "Request_ResourceNotFound":
                    case "ResourceNotFound":
                    case "ErrorItemNotFound":
                    case "itemNotFound":
                    case "ErrorInvalidUser":
                        // If picture not found, return the default image.
                        throw new Exception("ResourceNotFound");
                    case "TokenNotFound":
                        await httpContext.ChallengeAsync();
                        return null;
                    default:
                        return null;
                }
            }

            return pictureStream;
        }
        // Send an email message from the current user.
        public static async Task SendEmail(GraphServiceClient graphClient, IHostingEnvironment hostingEnvironment, string recipients, HttpContext httpContext)
        {
            if (recipients == null) return;

            var attachments = new MessageAttachmentsCollectionPage();

            try
            {
                // Load user's profile picture.
                var pictureStream = await GetMyPictureStream(graphClient, httpContext);

                // Copy stream to MemoryStream object so that it can be converted to byte array.
                var pictureMemoryStream = new MemoryStream();
                await pictureStream.CopyToAsync(pictureMemoryStream);

                // Convert stream to byte array and add as attachment.
                attachments.Add(new FileAttachment
                {
                    ODataType = "#microsoft.graph.fileAttachment",
                    ContentBytes = pictureMemoryStream.ToArray(),
                    ContentType = "image/png",
                    Name = "me.png"
                });
            }
            catch (Exception e)
            {
                switch (e.Message)
                {
                    case "ResourceNotFound":
                        break;
                    default:
                        throw;
                }
            }

            // Prepare the recipient list.
            var splitRecipientsString = recipients.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
            var recipientList = splitRecipientsString.Select(recipient => new Recipient
            {
                EmailAddress = new EmailAddress
                {
                    Address = recipient.Trim()
                }
            }).ToList();

            // Build the email message.
            var email = new Message
            {
                Body = new ItemBody
                {
                    Content = System.IO.File.ReadAllText(hostingEnvironment.WebRootPath + "/email_template.html"),
                    ContentType = BodyType.Html,
                },
                Subject = "Sent from the Microsoft Graph Connect sample",
                ToRecipients = recipientList,
                Attachments = attachments
            };

            await graphClient.Me.SendMail(email, true).Request().PostAsync();
        }

    }
}
