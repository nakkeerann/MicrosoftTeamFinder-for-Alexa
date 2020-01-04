// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.IO;
using Microsoft.Extensions.Logging;
using System.Net.Http;
using System.Net.Http.Headers;
using Microsoft.Azure.WebJobs;
using System.Collections.Concurrent;
using Newtonsoft.Json;

namespace Nav.AlexaTeamFinder
{
    public static class GraphHelper
    {
        private static string msGraphScope = "https://graph.microsoft.com/.default";
        private static string msGraphQuery = "";
        public static async Task<string> GetTeamsAsync(ILogger log, string teamSlot)
        {
            try
            {
                // Client creds from Azure AD APP which has permissions for accessing Microsoft teams data
                string clientId = ""; // Azure Client ID,
                string clientSecret = ""; // Copy the client secret
               
                var authority = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";
                authority = "https://login.microsoftonline.com/<tenant-name>.onmicrosoft.com/v2.0";
                IConfidentialClientApplication daemonClient = ConfidentialClientApplicationBuilder.Create(clientId)
                    .WithAuthority(authority)
                    .WithClientSecret(clientSecret)
                    .WithRedirectUri("https://navalexaskillapp.azurewebsites.net")
                    .Build();

                AuthenticationResult authResult = await daemonClient.AcquireTokenForClient(new[] { msGraphScope }).ExecuteAsync();
                log.LogInformation("Access Token: " + authResult.AccessToken);
                HttpClient client = new HttpClient();
                msGraphQuery = "https://graph.microsoft.com/beta/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team') and displayName eq '" + teamSlot + "'";
                HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, msGraphQuery);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
                HttpResponseMessage response = client.SendAsync(request).GetAwaiter().GetResult();
                
                string json = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                return json;
            }
            catch (Exception ex)
            {
                log.LogInformation("Graph Query Error" + ex.Message);
                return null;
            }
        }
    }
}