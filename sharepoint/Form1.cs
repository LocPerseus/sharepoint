using Azure.Identity;
using CsvHelper;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System.Configuration;
using System.Globalization;

namespace sharepoint
{
    public partial class Form1 : Form
    {
        private readonly GraphServiceClient _graphClient;
        private readonly Services _services;
        private List<Site> _sites;

        public Form1()
        {
            InitializeComponent();
            _graphClient = InitClient().Item1;
            _services = new Services(_graphClient, InitClient().Item2);
            _sites = new List<Site>();
        }

        private static (GraphServiceClient, ClientSecretCredential) InitClient()
        {
            var scopes = new[] { "https://graph.microsoft.com/.default" };

            // Multi-tenant apps can use "common",
            // single-tenant apps must use the tenant ID from the Azure portal
            var tenantId = ConfigurationManager.AppSettings["tenantId"];

            // Values from app registration
            var clientId = ConfigurationManager.AppSettings["clientId"];
            var clientSecret = ConfigurationManager.AppSettings["clientSecret"];

            // using Azure.Identity;
            var options = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };

            // https://learn.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
            var clientSecretCredential = new ClientSecretCredential(
                tenantId, clientId, clientSecret, options);


            return (new GraphServiceClient(clientSecretCredential, scopes), clientSecretCredential);
        }


        private async void Export_ClickAsync(object sender, EventArgs e)
        {
            string query = textBox1.Text;
            var sites = await _services.SearchSitesAsync(query);
            dataGridView1.DataSource = sites?.Value?.ToList();
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            _sites = new List<Site>();
            if (dataGridView1.SelectedRows.Count > 0)
            {
                // Get the selected rows
                var selectedRows = dataGridView1.SelectedRows;

                // Loop through the selected rows
                foreach (DataGridViewRow row in selectedRows)
                {
                    // Get the values of the cells in the selected row
                    var values = row.Cells.Cast<DataGridViewCell>().Select(cell => cell.Value).ToList();

                    // Do something with the selected row values
                    _sites.Add(new Site { Id = (string)values[22], DisplayName = (string)values[1] });
                }
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void ExportCSV_Click(object sender, EventArgs e)
        {
            var siteTasks = _sites.Select(async site =>
            {
                if (site.Id != null)
                {
                    var allLists = new List<Item>();
                    try
                    {
                        var listsInSite = await _services.GetListsInSite(site.Id);
                        if (listsInSite?.Value != null)
                        {
                            var fileTasks = listsInSite.Value.Select(async listSite =>
                            {
                                try
                                {
                                    var files = await _services.GetFiles(site.Id, listSite?.Id ?? "");
                                    if (files?.Value != null)
                                    {
                                        allLists.AddRange(files.Value);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    // Handle exception when calling _services.GetFiles
                                    MessageBox.Show(ex.Message);
                                }
                            });
                            await Task.WhenAll(fileTasks);
                        }
                    }
                    catch (Exception ex)
                    {
                        // Handle exception when calling _services.GetListsInSite
                        MessageBox.Show(ex.Message);
                    }

                    using var writer = new StreamWriter($"{site.DisplayName ?? site.Id}_{DateTime.Now:yyyyMMdd_hhmmss}.csv");
                    using var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);
                    csv.WriteRecords(allLists);

                }
            });

            if (_sites.Count == 0)
            {
                MessageBox.Show("Please select a site");
            }
            else
            {
                await Task.WhenAll(siteTasks);

                MessageBox.Show("Export Success!");
            }
        }
    }
}