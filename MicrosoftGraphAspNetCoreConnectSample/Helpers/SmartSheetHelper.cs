using MicrosoftGraphAspNetCoreConnectSample.Helpers.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace MicrosoftGraphAspNetCoreConnectSample.Helpers
{
    public static class SmartSheetHelper
    {
        public static async Task GetSheetShareInfo(string accessToken, string sheetId)
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

        public static async Task<string> GetPublishUrl(string accessToken, string sheetId)
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

        public static async Task<string> ObtainAccessToken(string url, string code, string clientId = "j6ex9cdw0ci3j9ucs2e", string clientSecret = "g8xirmd07iqr719i2f3")
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
