using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Identity.Client;
using WebApplication = Microsoft.AspNetCore.Builder.WebApplication;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

app.MapPost("/meeting", async () =>
{
    // 1. Configuration for the meeting
    // https://learn.microsoft.com/en-us/graph/api/resources/onlinemeeting?view=graph-rest-1.0
    var requestBody = new OnlineMeeting
    {
        StartDateTime = DateTime.Now,
        EndDateTime = DateTime.Now.AddYears(1),
        Subject = "Client Advisor Chat",

        AllowAttendeeToEnableCamera = false,
        AllowAttendeeToEnableMic = false,
        AllowMeetingChat = MeetingChatMode.Enabled,
        AllowParticipantsToChangeName = false,
        AllowTeamworkReactions = false,
        IsEntryExitAnnounced = true,
        ShareMeetingChatHistoryDefault = MeetingChatHistoryDefaultMode.All
    };

    // 2. Configuration for App Registration and graph client
    // TODO: Move to Key Vault
    var scopes = new[] { "https://graph.microsoft.com/.default" };

    var tenantId = "";
    var clientId = "";
    var clientSecret = "";
    var userId = "";

    var options = new ClientSecretCredentialOptions
    {
        AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
    };

    // https://learn.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
    var clientSecretCredential = new ClientSecretCredential(tenantId, clientId, clientSecret, options);
    var graphClient = new GraphServiceClient(clientSecretCredential, scopes);

    // 3. Create meeting on behalf of myself
    var result = await graphClient.Users[userId].OnlineMeetings.PostAsync(requestBody);
    Console.WriteLine($"Meeting result: {result.Id}");

    // TODO: Save resultant meeting details to database

    // 4. Send chat message

    var delegatedClient = GraphServiceClientBuilder.GetDelegatedClient();

    var chatThreadId = result.ChatInfo.ThreadId;
    var chatMessage = new ChatMessage
    {
        Body = new ItemBody()
        {
            Content = "This is an automated message"
        }
    };
    var chatResult = await delegatedClient.Me.Chats[chatThreadId].Messages.PostAsync(chatMessage);
    Console.WriteLine($"Chat chatResult: {chatResult.Id}");
})
.WithName("CreateMeeting")
.WithOpenApi();

app.Run();

public class GraphServiceClientBuilder
{
    public static GraphServiceClient GetDelegatedClient()
    {
        var scopes = new[] { "https://graph.microsoft.com/.default" };

        var tenantId = "";
        var clientId = "";
        var clientSecret = "";

        var options = new OnBehalfOfCredentialOptions
        {
            AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
        };

        // This object will cache tokens in-memory - keep it as a singleton
        var app = ConfidentialClientApplicationBuilder.Create(clientId)
                // Don't specify authority here, we'll do it on the request 
                .WithClientSecret(clientSecret)
                .Build();

        // If instead you need to re-create the ConfidentialClientApplication on each request, you MUST customize 
        // the cache serialization (see below)

        // When making the request, specify the tenant-based authority
        var authResult = app.AcquireTokenForClient(scopes: scopes)        // Uses the token cache automatically, which is optimized for multi-tenant access
                .WithAuthority(AzureCloudInstance.AzurePublic, tenantId)  // Do not use "common" or "organizations"!
                .ExecuteAsync().Result;

        // This is the incoming token to exchange using on-behalf-of flow
        var oboToken = authResult.AccessToken;

        var onBehalfOfCredential = new OnBehalfOfCredential(tenantId, clientId, clientSecret, oboToken, options);
        var graphClient = new GraphServiceClient(onBehalfOfCredential, scopes);

        return graphClient;
    }
}