using System.Threading.Tasks;
using Alexa.NET.Request;
using Alexa.NET.Request.Type;
using Alexa.NET;
using Alexa.NET.Response;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Xml;
using System.Linq;
using System;

namespace Nav.AlexaTeamFinder
{
    public static class AlexaTeamFinder
    {
        [FunctionName("AlexaTeamFinder")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            string json = await req.ReadAsStringAsync();
            var skillRequest = JsonConvert.DeserializeObject<SkillRequest>(json);
            var requestType = skillRequest.GetRequestType();
            
            SkillResponse response = null;

            if (requestType == typeof(LaunchRequest))
            {
                response = ResponseBuilder.Tell("Hello, Welcome to the Team Finder skill!");
                log.LogInformation("Welcome Message Announced");
                response.Response.ShouldEndSession = false;
            }
            else if (requestType == typeof(IntentRequest))
            {
                var intentRequest = skillRequest.Request as IntentRequest;

                if (intentRequest.Intent.Name == "teaminformation")
                {
                    var teamSlots = intentRequest.Intent.Slots["teamname"].Value;
                    log.LogInformation("Finding Team Information");
                    var teamsStr = await GraphHelper.GetTeamsAsync(log,teamSlots);
                    JObject teamsJson = JObject.Parse(teamsStr);
                    var teamsValue = teamsJson["value"];
                    var message ="";
                    int counter = 1;
                    foreach(var team in teamsJson["value"]){
                        message += "Number "+ counter+", Team Name: "+team["displayName"];
                        message += ", Description: "+ team["description"]+";";
                        counter++;
                    }
                    response = ResponseBuilder.Tell(message);                    
                }                
            }
            else if (requestType == typeof(SessionEndedRequest))
            {
                log.LogInformation("Session ended");
                response = ResponseBuilder.Empty();
                response.Response.ShouldEndSession = true;
            }
            return new OkObjectResult(response);
        }
    }
}
