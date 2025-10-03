using Azure.Identity;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace MsalExample;

public class MsGraphExamplesUsingManagedIdentity
{
    private readonly ILogger<MsGraphExamplesUsingManagedIdentity> _logger;
    private readonly IConfiguration _configuration;
    private readonly GraphServiceClient _graphServiceClient;

    public MsGraphExamplesUsingManagedIdentity(
        ILogger<MsGraphExamplesUsingManagedIdentity> logger,
        IConfiguration configuration)
    {
        _logger = logger;
        _configuration = configuration;
        var credential = new DefaultAzureCredential();
        var scopes = new[] { "https://graph.microsoft.com/.default" };
        _graphServiceClient = new GraphServiceClient(credential, scopes);
    }

    [Function("miusers")]
    public async Task<IActionResult> GetUsersAsync(
        [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "mi/users/{username}")] HttpRequest req,
        string username)
    {
        var user = await _graphServiceClient.Users[username].GetAsync(config =>
        {
            config.QueryParameters.Select = new[] { "id", "displayName", "userPrincipalName" };
        });

        var userModel = new UserModel(user?.Id, user?.DisplayName, user?.UserPrincipalName);
        return new OkObjectResult(userModel);
    }

    [Function("miallusers")]
    public async Task<IActionResult> GetAllUsersAsync(
        [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "mi/allusers")] HttpRequest req)
    {
        var allUsers = new List<UserModel>();
        var result = await _graphServiceClient.Users.GetAsync(config =>
        {
            config.QueryParameters.Select = new[] { "id", "displayName", "userPrincipalName" };
        });

        AddUsersFromResult(result, allUsers);

        while (result?.OdataNextLink != null)
        {
            result = await _graphServiceClient.Users.WithUrl(result.OdataNextLink).GetAsync();
            AddUsersFromResult(result, allUsers);
        }

        return new OkObjectResult(allUsers);
    }

    [Function("miusergroups")]
    public async Task<IActionResult> GetUserGroupsAsync(
        [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "mi/usergroups/{username}")] HttpRequest req,
        string username)
    {
        var allGroups = new List<GroupModel>();
        var result = await _graphServiceClient.Users[username].MemberOf.GraphGroup.GetAsync(config =>
        {
            config.QueryParameters.Select = new[] { "id", "displayName", "description" };
        });

        AddGroupsFromResult(result, allGroups);

        while (result?.OdataNextLink != null)
        {
            result = await _graphServiceClient.Groups.WithUrl(result.OdataNextLink).GetAsync();
            AddGroupsFromResult(result, allGroups);
        }

        return new OkObjectResult(allGroups);
    }

    [Function("migroups")]
    public async Task<IActionResult> GetGroupsAsync(
        [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "mi/groups")] HttpRequest req)
    {
        var allGroups = new List<GroupModel>();

        try
        {
            var groupsPage = await _graphServiceClient.Groups.GetAsync();
            AddGroupsFromResult(groupsPage, allGroups);

            while (groupsPage?.OdataNextLink != null)
            {
                groupsPage = await _graphServiceClient.Groups.WithUrl(groupsPage.OdataNextLink).GetAsync();
                AddGroupsFromResult(groupsPage, allGroups);
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error getting groups");
        }

        return new OkObjectResult(allGroups);
    }

    [Function("miusertransitivegroups")]
    public async Task<IActionResult> GetUserTransitiveGroupsAsync(
        [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "mi/usertransitivegroups/{username}")] HttpRequest req,
        string username)
    {
        var allGroups = new List<GroupModel>();
        var result = await _graphServiceClient.Users[username].TransitiveMemberOf.GraphGroup.GetAsync(config =>
        {
            config.QueryParameters.Select = new[] { "id", "displayName", "description" };
        });

        AddGroupsFromResult(result, allGroups);

        while (result?.OdataNextLink != null)
        {
            result = await _graphServiceClient.Users[username].TransitiveMemberOf.GraphGroup.WithUrl(result.OdataNextLink).GetAsync();
            AddGroupsFromResult(result, allGroups);
        }

        return new OkObjectResult(allGroups);
    }

    private static void AddUsersFromResult(UserCollectionResponse? result, List<UserModel> allUsers)
    {
        if (result?.Value != null)
        {
            allUsers.AddRange(result.Value.Select(user => new UserModel(user?.Id, user?.DisplayName, user?.UserPrincipalName)));
        }
    }

    private static void AddGroupsFromResult(GroupCollectionResponse? result, List<GroupModel> allGroups)
    {
        if (result?.Value != null)
        {
            allGroups.AddRange(result.Value.Select(x => new GroupModel(x.Id, x.DisplayName, x.Description)));
        }
    }
}

