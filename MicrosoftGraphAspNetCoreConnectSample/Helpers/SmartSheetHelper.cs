using MicrosoftGraphAspNetCoreConnectSample.Helpers.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace MicrosoftGraphAspNetCoreConnectSample.Helpers
{
    public static class SmartSheetHelper
    {
        public static async Task<SmartSheetShare> GetSheetShareInfo(string accessToken, string sheetId)
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

            SmartSheetShare shareInfo = null;
            if (response.IsSuccessStatusCode)
            {
                reponseUrl = await response.Content.ReadAsStringAsync();
                shareInfo = JsonConvert.DeserializeObject<SmartSheetShare>(reponseUrl);

                Debug.WriteLine("Shared with:" + shareInfo.Data.Count);
            }
            return shareInfo;
        }

        public static async Task<string> GetPublishUrl(string accessToken, string sheetId)
        {
            string responseUrl = string.Empty;
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
                string apiResponse = await response.Content.ReadAsStringAsync();
                Dictionary<string, string> responseInfo = JsonConvert.DeserializeObject<Dictionary<string, string>>(apiResponse);
                if (responseInfo["readWriteEnabled"].Equals("true", StringComparison.InvariantCultureIgnoreCase))
                {
                    responseUrl = responseInfo["readWriteUrl"];
                }
            }

            if (string.IsNullOrWhiteSpace(responseUrl))
            {
                responseUrl = await PublishRWSheet(accessToken, sheetId);
            }

            return responseUrl;
        }

        private static async Task<string> PublishRWSheet(string accessToken, string sheetId)
        {
            string returnInfo = string.Empty;
            string url = string.Empty;
            if (!string.IsNullOrWhiteSpace(sheetId))
            {
                url = $"https://api.smartsheet.com/2.0/sheets/{sheetId}/publish ";
            }

            if (string.IsNullOrWhiteSpace(accessToken))
            {
                throw new Exception("Provided Smartsheet Code cannot be null");
            }

            HttpClient smartSheetApiClient = new HttpClient();
            smartSheetApiClient.DefaultRequestHeaders.Add("Authorization", "Bearer " + accessToken);
            // smartSheetApiClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            SheetPublish publishInfo = new SheetPublish() { IcalEnabled = false, ReadOnlyFullEnabled = false, ReadOnlyLiteEnabled = false, ReadWriteEnabled = true, };
            var response = await smartSheetApiClient.PutAsync(url, new StringContent(JsonConvert.SerializeObject(publishInfo), Encoding.UTF8, "application/json"));

            if (response.IsSuccessStatusCode)
            {
                string responseUrl = await response.Content.ReadAsStringAsync();
                SheetPublishResponse responseInfo = JsonConvert.DeserializeObject<SheetPublishResponse>(responseUrl);
                if (responseInfo.ResultCode == 0)
                {
                    returnInfo = responseInfo.Result.ReadWriteUrl;
                }
            }

            return returnInfo;
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

    #region Smartsheet model helpers
    public class SheetPublish
    {
        [JsonProperty("readOnlyLiteEnabled")]
        public bool ReadOnlyLiteEnabled { get; set; }

        [JsonProperty("readOnlyFullEnabled")]
        public bool ReadOnlyFullEnabled { get; set; }

        [JsonProperty("readWriteEnabled")]
        public bool ReadWriteEnabled { get; set; }

        [JsonProperty("icalEnabled")]
        public bool IcalEnabled { get; set; }
    }

    public class SheetPublishResponse
    {
        [JsonProperty("message")]
        public string Message { get; set; }

        [JsonProperty("resultCode")]
        public long ResultCode { get; set; }

        [JsonProperty("result")]
        public Result Result { get; set; }
    }

    public class Result
    {
        [JsonProperty("icalEnabled")]
        public bool IcalEnabled { get; set; }

        [JsonProperty("readOnlyFullEnabled")]
        public bool ReadOnlyFullEnabled { get; set; }

        [JsonProperty("readOnlyLiteEnabled")]
        public bool ReadOnlyLiteEnabled { get; set; }

        [JsonProperty("readOnlyLiteUrl")]
        public Uri ReadOnlyLiteUrl { get; set; }

        [JsonProperty("readWriteUrl")]
        public string ReadWriteUrl { get; set; }

        [JsonProperty("readWriteEnabled")]
        public bool ReadWriteEnabled { get; set; }
    }
    #endregion

    public class SheetInformation
    {
        public string SheetName { get; set; }

        public string SheetId { get; set; }

        public string SheetRWUrl { get; set; }

        public List<string> Collaborators { get; set; }

        public string TeamsUrl { get; set; }
    }

}
