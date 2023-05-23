using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System.Net.Http.Headers;
using System.Net.Http.Json;

namespace sharepoint
{
    public class Services
    {
        private readonly GraphServiceClient _graphServiceClient;
        private readonly ClientSecretCredential _clientSecretCredential;
        private string accessToken = "";

        public Services(GraphServiceClient graphService, ClientSecretCredential clientSecretCredential)
        {
            _graphServiceClient = graphService;
            _clientSecretCredential = clientSecretCredential;
            _ = GetToken();
        }

        public async Task GetToken()
        {
            var tokenRequestContext = new TokenRequestContext(new[] { "https://graph.microsoft.com/.default" });
            var token = await _clientSecretCredential.GetTokenAsync(tokenRequestContext);
            accessToken = token.Token;
        }

        public async Task<SiteCollectionResponse?> SearchSitesAsync(string query)
        {
            return await _graphServiceClient.Sites.GetAsync((requestConfiguration) =>
            {
                requestConfiguration.QueryParameters.Search = query;
            });
        }

        public async Task<ListCollectionResponse?> GetListsInSite(string siteId)
        {
            return await _graphServiceClient.Sites[siteId].Lists.GetAsync();
        }

        public async Task<ItemResponse?> GetFiles(string siteId, string listId)
        {
            using var httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            httpClient.DefaultRequestHeaders.Add("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");

            var response = await httpClient.GetAsync($"https://graph.microsoft.com/v1.0/sites/{siteId}/lists/{listId}/items?$expand=fields,driveItem&$filter=fields/ContentType eq 'Document'");

            if (response.IsSuccessStatusCode)
            {
                var content = await response.Content.ReadFromJsonAsync<ItemResponse>();
                // Handle the response content
                return content;
            }
            else
            {
                // Handle the error
                MessageBox.Show(response.ReasonPhrase);
                return null;
            }
        }
    }
}