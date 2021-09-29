using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace PresenceNotificationsPoc.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class SubscriptionController : ControllerBase
    {
        private const string SharedKey = "Blablabla";

        private readonly ILogger<SubscriptionController> _logger;
        private readonly GraphConfig _graphConfig;
        private readonly IConfiguration _configuration;

        private static Dictionary<string, string> Users = new();

        public SubscriptionController(ILogger<SubscriptionController> logger, GraphConfig graphConfig,
            IConfiguration configuration)
        {
            _logger = logger;
            _graphConfig = graphConfig;
            _configuration = configuration;
        }

        [HttpGet]
        public async Task<IActionResult> Create(string userPrincipalName = null)
        {
            var httpClient = new HttpClient()
            {
                BaseAddress = new Uri("https://graph.microsoft.com/"),
            };
            httpClient.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", _graphConfig.AccessToken);

            await RemoveExistingSubscriptions(httpClient);

            var userIds = new Dictionary<string, string>();

            if (userPrincipalName != null)
            {
                var userId = await GetUserId(httpClient, userPrincipalName);
                userIds.Add(userId, userPrincipalName);
            }

            var users = await GetAllUsers(httpClient);

            foreach (var user in users)
            {
                if (!userIds.ContainsKey(user.Id))
                    userIds.Add(user.Id, user.UserPrincipalName);
            }

            var userIdsString = string.Join(',', userIds.Keys.Take(650).Select(userId => $"'{Uri.EscapeDataString(userId)}'"));

            var request = new SubscriptionRequest
            {
                changeType = "updated",
                notificationUrl = _graphConfig.ExternalBaseUrl + "/subscription",
                resource = $"/communications/presences?$filter=id in ({userIdsString})",
                expirationDateTime = DateTimeOffset.UtcNow.AddDays(1),
                clientState = SharedKey,
            };
            var response = await httpClient.PostAsync("v1.0/subscriptions",
                new StringContent(JsonSerializer.Serialize(request), Encoding.UTF8, "application/json"));

            if (!response.IsSuccessStatusCode)
            {
                var body = await response.Content.ReadAsStringAsync();
                return BadRequest(new
                {
                    ResponseCode = response.StatusCode,
                    ResponseBody = body
                });
            }
            else
            {
                Users = userIds;
                return Ok("Subscription successful!");
            }
        }

        private async Task RemoveExistingSubscriptions(HttpClient httpClient)
        {
            var response = await httpClient.GetAsync("v1.0/subscriptions");
            response.EnsureSuccessStatusCode();
            var responseBody = await response.Content.ReadAsStringAsync();
            var responseDoc = JsonDocument.Parse(responseBody);

            foreach (var item in responseDoc.RootElement.GetProperty("value").EnumerateArray())
            {
                string resource = item.GetProperty("resource").GetString();
                if (resource.StartsWith("/communications/presences?$filter="))
                {
                    string itemId = item.GetProperty("id").GetString();
                    _logger.LogDebug("Removing subscription with id {SubscriptionId}", itemId);
                    var delResponse = await httpClient.DeleteAsync($"v1.0/subscriptions/{Uri.EscapeUriString(itemId)}");
                    delResponse.EnsureSuccessStatusCode();
                    _logger.LogDebug("Removed subscription with id {SubscriptionId}", itemId);
                }
            }
        }

        [HttpPost()]
        public async Task<IActionResult> Updates([FromQuery] string validationToken = null)
        {
            if (validationToken != null)
            {
                _logger.LogInformation("Subscription confirmed");
                Response.ContentType = "text/plain";
                return Ok(validationToken);
            }

            try
            {
                // Read the body
                using var reader = new StreamReader(Request.Body);
                var jsonPayload = await reader.ReadToEndAsync();

                // Use the Graph client's serializer to deserialize the body
                var response = JsonSerializer.Deserialize<PresenceSubscriptionResponse>(jsonPayload);
                foreach (var resourceUpdate in response.PresenceResourceUpdates)
                {
                    Users.TryGetValue(resourceUpdate.PresenceUpdate?.Id, out string upn);
                    _logger.LogInformation("Id: {Id}, Upn: {Upn}, Availability: {Availability}",
                        resourceUpdate.PresenceUpdate?.Id, upn, resourceUpdate.PresenceUpdate?.Availability);
                }

                return Ok(response);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to process update body");
                return new StatusCodeResult(500);
            }
        }

        private async Task<string> GetUserId(HttpClient httpClient, string userPrincipalName)
        {
            var response =
                await httpClient.GetAsync(
                    $"v1.0/users/{Uri.EscapeUriString(userPrincipalName)}?$select=id,accountEnabled");
            if (response.StatusCode == System.Net.HttpStatusCode.NotFound) return null;
            response.EnsureSuccessStatusCode();

            var jsonPayload = await response.Content.ReadAsStringAsync();
            var doc = JsonDocument.Parse(jsonPayload);
            if (doc.RootElement.GetProperty("accountEnabled").GetBoolean())
                return doc.RootElement.GetProperty("id").GetString();

            return null;
        }

        private async Task<List<User>> GetAllUsers(HttpClient httpClient, string url = @"v1.0/users?$filter=accountEnabled eq true&$select=id,userPrincipalName", List<User> users = null)
        {
            var response =
                await httpClient.GetAsync(url);

            var jsonPayload = await response.Content.ReadAsStringAsync();
            var doc = JsonDocument.Parse(jsonPayload);

            string nextUrl = (doc.RootElement.TryGetProperty("@odata.nextLink", out JsonElement nextUrlJson)) ? nextUrlJson.GetString() : null;

            var result = users ?? new List<User>();
            result.AddRange(doc.RootElement.GetProperty("value")
                .EnumerateArray()
                .Select(userJson => new Controllers.User
                {
                    Id = userJson.GetProperty("id").GetString(),
                    UserPrincipalName = userJson.GetProperty("userPrincipalName").GetString(),
                }));

            if (nextUrl != null)
            {
                return await  GetAllUsers(httpClient, nextUrl, result);
            }

            return result;
        }
    }

    public class User
    {
        public string Id { get; set; }
        public string UserPrincipalName { get; set; }
    }

    public class PresenceSubscriptionResponse
    {
        [JsonPropertyName("value")] public PresenceResourceUpdate[] PresenceResourceUpdates { get; set; }

        public class PresenceResourceUpdate
        {
            [JsonPropertyName("resourceData")] public PresenceUpdate PresenceUpdate { get; set; }

            [JsonPropertyName("subscriptionId")] public string SubscriptionId { get; set; }

            [JsonPropertyName("subscriptionExpirationDateTime")]
            public DateTimeOffset SubscriptionExpirationDateTime { get; set; }

            [JsonPropertyName("tenantId")] public string TenantId { get; set; }
        }

        public class PresenceUpdate
        {
            [JsonPropertyName("id")] public string Id { get; set; }
            [JsonPropertyName("availability")] public string Availability { get; set; }
            [JsonPropertyName("activity")] public string Activity { get; set; }
        }
    }

    public class SubscriptionRequest
    {
        public string changeType { get; set; }
        public string notificationUrl { get; set; }
        public string resource { get; set; }
        public DateTimeOffset expirationDateTime { get; set; }
        public string clientState { get; set; }
    }
}