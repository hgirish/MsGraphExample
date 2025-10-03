using Azure.Identity;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System.Text.Json;

namespace MsalExample;

public class MsGraphExamples
{
    private readonly ILogger<MsGraphExamples> _logger;
    private readonly IConfiguration _configuration;
    private readonly GraphServiceClient _graphServiceClient;
    public MsGraphExamples(ILogger<MsGraphExamples> logger, IConfiguration configuration)
    {
        _logger = logger;
        _configuration = configuration;
        var options = new TokenCredentialOptions
        {
            AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
        };
        // Example using ClientSecretCredential
        var scopes = new[] { "https://graph.microsoft.com/.default" };
        var tenantId = configuration["AZURE_TENANT_ID"];
        var clientId = configuration["AZURE_CLIENT_ID"];
        var clientSecret = configuration["AZURE_CLIENT_SECRET"];

        var clientSecretCredential = new ClientSecretCredential(tenantId, clientId, clientSecret, options);

        _graphServiceClient = new GraphServiceClient(clientSecretCredential, scopes);

    }

    [Function("users")]
    public async Task<IActionResult> GetUsersAsync([HttpTrigger(
        AuthorizationLevel.Anonymous, "get", Route ="users/{username}")] HttpRequest req,
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
    [Function("usergroups")]
    public async Task<IActionResult> GetUserGroupsAsync(
        [HttpTrigger(
        AuthorizationLevel.Anonymous, "get", Route ="usergroups/{username}")] HttpRequest req,
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
    [Function("groups")]
    public async Task<IActionResult> GetGroupsAsync(
         [HttpTrigger(
        AuthorizationLevel.Anonymous, "get", Route ="groups")] HttpRequest req)
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

    [Function("usertransitivegroups")]
    public async Task<IActionResult> GetUserTransitiveGroupsAsync(
          [HttpTrigger(
        AuthorizationLevel.Anonymous, "get",Route ="usertransitivegroups/{username}")] HttpRequest req,
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
public record GroupModel(string? Id, string? DisplayName, string? Description);

public record UserModel(string? Id, string? DisplayName, string? UserPrincipalName);
