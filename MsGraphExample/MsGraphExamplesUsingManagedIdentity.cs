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
        var clientId = _configuration["MSAL_CLIENT_ID"];
        var credential = new ManagedIdentityCredential(clientId);

        //var credential = new DefaultAzureCredential(manage) ;
        //DefaultAzureCredential credential = new DefaultAzureCredentialBuilder()
        //    .managedIdentityClientId("your-user-assigned-managed-identity-client-id")
        //    .build();
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

        var userModel = user.ToUserModel();
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

        result.AddUsersTo(allUsers);

        while (result?.OdataNextLink != null)
        {
            result = await _graphServiceClient.Users.WithUrl(result.OdataNextLink).GetAsync();
            result.AddUsersTo(allUsers);
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

        result.AddGroupsTo(allGroups);

        while (result?.OdataNextLink != null)
        {
            result = await _graphServiceClient.Groups.WithUrl(result.OdataNextLink).GetAsync();
            result.AddGroupsTo(allGroups);
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
            groupsPage.AddGroupsTo(allGroups);

            while (groupsPage?.OdataNextLink != null)
            {
                groupsPage = await _graphServiceClient.Groups.WithUrl(groupsPage.OdataNextLink).GetAsync();
                groupsPage.AddGroupsTo(allGroups);
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

        result.AddGroupsTo(allGroups);

        while (result?.OdataNextLink != null)
        {
            result = await _graphServiceClient.Users[username].TransitiveMemberOf.GraphGroup.WithUrl(result.OdataNextLink).GetAsync();
            result.AddGroupsTo(allGroups);
        }

        return new OkObjectResult(allGroups);
    }
}

public static class UserCollectionResponseExtensions
{
    public static void AddUsersTo(this UserCollectionResponse? result, List<UserModel> allUsers)
    {
        if (result?.Value != null)
        {
            allUsers.AddRange(result.Value.Select(user => user.ToUserModel()));
        }
    }

    public static UserModel ToUserModel(this User? user)
    {
        return new UserModel(user?.Id, user?.DisplayName, user?.UserPrincipalName);
    }
}

public static class GroupCollectionResponseExtensions
{
    public static void AddGroupsTo(this GroupCollectionResponse? result, List<GroupModel> allGroups)
    {
        if (result?.Value != null)
        {
            allGroups.AddRange(result.Value.Select(x => x.ToGroupModel()));
        }
    }

    public static GroupModel ToGroupModel(this Group? group)
    {
        return new GroupModel(group?.Id, group?.DisplayName, group?.Description);
    }
}

