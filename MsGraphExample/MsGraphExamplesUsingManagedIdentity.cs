using Azure.Identity;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System.Security.Claims;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using static Microsoft.ApplicationInsights.MetricDimensionNames.TelemetryContext;

namespace MsalExample;

public class MsGraphExamplesUsingManagedIdentity
{
    private readonly ILogger<MsGraphExamplesUsingManagedIdentity> _logger;
    private readonly IConfiguration _configuration;
    private readonly GraphServiceClient _graphServiceClient;
    public MsGraphExamplesUsingManagedIdentity(ILogger<MsGraphExamplesUsingManagedIdentity> logger, IConfiguration configuration)
    {
        _logger = logger;
        _configuration = configuration;
        var options = new TokenCredentialOptions
        {
            AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
            
        };
        // Example using ClientSecretCredential
        var scopes = new[] { "https://graph.microsoft.com/.default" };
  
        var credential = new DefaultAzureCredential();
        _graphServiceClient = new GraphServiceClient(credential, scopes);

    }

    [Function("miusers")]
    public async Task<IActionResult> GetUsersAsync([HttpTrigger(
        AuthorizationLevel.Anonymous, "get", Route ="mi/users/{username}")] HttpRequest req,
        string username)
    {
        var user = await _graphServiceClient.Users[username].GetAsync((requestConfiguration) =>
        {
            // requestConfiguration.QueryParameters.Filter = "appRoleAssignments/$count gt 0";
            requestConfiguration.QueryParameters.Select = new string[] { "id", "displayName", "userPrincipalName" };
        });
        var userModel = new UserModel(

             user?.Id,
            user?.DisplayName,
            user?.UserPrincipalName
        );
        //var user = await graphClient.Users.GetAsync();
        var json = JsonSerializer.Serialize(userModel);

        return new OkObjectResult(userModel);
    }
    [Function("miallusers")]
    public async Task<IActionResult> GetAllUsersAsync([HttpTrigger(
        AuthorizationLevel.Anonymous, "get", Route ="mi/allusers")] HttpRequest req
     )
    {
        var allUsers = new List<UserModel>();
        var result  = await _graphServiceClient.Users.GetAsync((requestConfiguration) =>
        {
            // requestConfiguration.QueryParameters.Filter = "appRoleAssignments/$count gt 0";
            requestConfiguration.QueryParameters.Select = new string[] { "id", "displayName", "userPrincipalName" };
        });
       
        if (result?.Value != null)
        {
            allUsers.AddRange(result.Value.Select(user => new UserModel(user?.Id, user?.DisplayName, user?.UserPrincipalName)));
        }
        while (result?.OdataNextLink != null)
        {
            result = await _graphServiceClient.Users.WithUrl(result?.OdataNextLink).GetAsync();
            

            if (result?.Value != null)
            {
                allUsers.AddRange(result.Value.Select(user => new UserModel(user?.Id, user?.DisplayName, user?.UserPrincipalName)));
            }
        }
        //var user = await graphClient.Users.GetAsync();
        var json = JsonSerializer.Serialize(allUsers);

        return new OkObjectResult(json);
    }
    [Function("miusergroups")]
    public async Task<IActionResult> GetUserGroupsAsync(
        [HttpTrigger(
        AuthorizationLevel.Anonymous, "get", Route ="mi/usergroups/{username}")] HttpRequest req,
        string username)
    {
        //// Get direct memberships for a specific user by ID
        //var directMemberships = await _graphServiceClient.Users[username]
        //    .MemberOf           
        //    .GetAsync();

        // Or for the currently signed-in user (delegated scenario)
        // var directMemberships = await client.Me
        //     .MemberOf
        //     .Request()
        //     .GetAsync();
        var allGroups = new List<GroupModel>();

        var result = await _graphServiceClient.Users[username].MemberOf.GraphGroup.GetAsync((requestConfiguration) =>
        {
            // requestConfiguration.QueryParameters.Filter = "appRoleAssignments/$count gt 0";
            requestConfiguration.QueryParameters.Select = new string[] { "id", "displayName", "description" };
        });
        if (result?.Value != null)
        {
            allGroups.AddRange(result.Value.Select(x => new GroupModel
           (
                         x.Id,
                        x.DisplayName,
                        x.Description
                    )));
            while (result?.OdataNextLink != null)
            {
                result = await _graphServiceClient.Groups.WithUrl(result.OdataNextLink).GetAsync();
                if (result?.Value != null)
                {
                    allGroups.AddRange(result.Value.Select(x => new GroupModel
                   (
                         x.Id,
                        x.DisplayName,
                        x.Description
                    )));
                }
            }
        }
        return new OkObjectResult(allGroups);
    }
    [Function("migroups")]
    public async Task<IActionResult> GetGroupsAsync(
         [HttpTrigger(
        AuthorizationLevel.Anonymous, "get", Route ="mi/groups")] HttpRequest req)
    {
        var allGroups = new List<GroupModel>();

        try
        {
            // Get the first page of groups
            var groupsPage = await _graphServiceClient.Groups.GetAsync();

            if (groupsPage?.Value != null)
            {
                allGroups.AddRange(groupsPage.Value.Select(x => new GroupModel
                (
                         x.Id,
                        x.DisplayName,
                        x.Description
                    )));

                // Handle pagination if there are more groups
                while (groupsPage?.OdataNextLink != null)
                {
                    groupsPage = await _graphServiceClient.Groups.WithUrl(groupsPage.OdataNextLink).GetAsync();
                    if (groupsPage?.Value != null)
                    {
                        allGroups.AddRange(groupsPage.Value.Select(x => new GroupModel
                        (
                         x.Id,
                        x.DisplayName,
                        x.Description
                    )));
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting groups: {ex.Message}");
        }

        return new OkObjectResult(allGroups);
    }

    [Function("miusertransitivegroups")]
    public async Task<IActionResult> GetUserTransitiveGroupsAsync(
          [HttpTrigger(
        AuthorizationLevel.Anonymous, "get",Route ="mi/usertransitivegroups/{username}")] HttpRequest req,
          string username)
    {

        var allGroups = new List<GroupModel>();

        var result = await _graphServiceClient.Users[username].TransitiveMemberOf.GraphGroup.GetAsync((requestConfiguration) =>
        {
            // requestConfiguration.QueryParameters.Filter = "appRoleAssignments/$count gt 0";
            requestConfiguration.QueryParameters.Select = new string[] { "id", "displayName", "description" };
        });
        if (result?.Value != null)
        {
            allGroups.AddRange(result.Value.Select(x => new GroupModel
            (
                         x.Id,
                        x.DisplayName,
                        x.Description
                    )));
            while (result?.OdataNextLink != null)
            {
                result = await _graphServiceClient.Users[username].TransitiveMemberOf.GraphGroup.WithUrl(result.OdataNextLink).GetAsync();
                if (result?.Value != null)
                {
                    allGroups.AddRange(result.Value.Select(x => new GroupModel
                    (
                         x.Id,
                        x.DisplayName,
                        x.Description
                    )));
                }
            }
        }
        return new OkObjectResult(allGroups);
    }
}

