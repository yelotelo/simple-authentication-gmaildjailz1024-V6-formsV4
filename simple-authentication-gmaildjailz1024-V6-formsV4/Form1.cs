using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace simple_authentication_gmaildjailz1024_V6_formsV4
{
    public partial class Form1 : Form
    {
        private Button btnMe;
        private Button btnUsers;
        private Button btnApplications;
        private Button btnGroups;
        private Button btnSignIn;
        private TextBox txtOutput;
        private ListBox lstUsers; // NEW: list of users

        public Form1()
        {
            InitializeComponent();

            // create controls dynamically so no designer changes are required
            var panel = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 40, AutoSize = false };


            btnSignIn = new Button { Text = "Sign In", Width = 120, Height = 34, Margin = new Padding(6) };
            btnMe = new Button { Text = "Me", Width = 120, Height = 34, Margin = new Padding(6) };
            btnUsers = new Button { Text = "Users", Width = 120, Height = 34, Margin = new Padding(6) };
            btnApplications = new Button { Text = "Applications", Width = 120, Height = 34, Margin = new Padding(6) };
            btnGroups = new Button { Text = "Groups", Width = 120, Height = 34, Margin = new Padding(6) };

            panel.Controls.Add(btnMe);
            panel.Controls.Add(btnUsers);
            panel.Controls.Add(btnApplications);
            panel.Controls.Add(btnGroups);
            panel.Controls.Add(btnSignIn);

            // Create a split layout: left = output, right = users list
            var split = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Vertical, SplitterDistance = (int)(Width * 0.6) };

            txtOutput = new TextBox { Multiline = true, Dock = DockStyle.Fill, ScrollBars = ScrollBars.Both, ReadOnly = true };

            lstUsers = new ListBox { Dock = DockStyle.Fill };
            lstUsers.SelectedIndexChanged += LstUsers_SelectedIndexChanged; // handle selection

            split.Panel1.Controls.Add(txtOutput);
            split.Panel2.Controls.Add(lstUsers);

            Controls.Add(split);
            Controls.Add(panel);

            btnMe.Click += BtnMe_Click;
            btnUsers.Click += BtnUsers_Click;
            btnApplications.Click += BtnApplications_Click;
            btnGroups.Click += BtnGroups_Click;
            btnSignIn.Click += BtnSignIn_Click;

            // Start with action buttons disabled until explicit sign-in
            btnMe.Enabled = btnUsers.Enabled = btnApplications.Enabled = btnGroups.Enabled = false;

            // If a token was persisted from a previous session, load it and enable actions
            LoadCachedTokenFromFile();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        // Ensure we have an access token. This is called automatically by the endpoint buttons if needed.
        private async Task EnsureTokenAsync()
        {
            // disable action buttons while acquiring
            btnMe.Enabled = btnUsers.Enabled = btnApplications.Enabled = btnGroups.Enabled = false;
            txtOutput.Clear();

            try
            {
                string tenant = "djailz1024gmail.onmicrosoft.com";
                string clientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e";

                string authority = $"https://login.microsoftonline.com/{tenant}";

                string[] scopes = new[]
                {
                    "offline_access",
                    "group.read.all",
                    "openid",
                    "user.read.all"
                };

                var app = PublicClientApplicationBuilder
                    .Create(clientId)
                    .WithAuthority(authority)
                    .WithRedirectUri("http://localhost")
                    .Build();

                AppendLine("Acquiring token interactively...");
                var result = await app.AcquireTokenInteractive(scopes).ExecuteAsync();

                string token = result.AccessToken;
                AppendHeader("TOKEN");
                AppendLine(token);

                // Persist token to file and to a field so other buttons can use it
                File.WriteAllText("Token-v9.txt", token);
                _cachedToken = token;

                AppendLine("Token acquired.");
            }
            catch (MsalException mex)
            {
                AppendLine($"MSAL error: {mex.Message}");
            }
            catch (HttpRequestException hex)
            {
                AppendLine($"HTTP error: {hex.Message}");
            }
            catch (Exception ex)
            {
                AppendLine($"Unexpected error: {ex.Message}");
            }
            finally
            {
                btnMe.Enabled = btnUsers.Enabled = btnApplications.Enabled = btnGroups.Enabled = true;
            }
        }

        // Cached token for the separate action buttons
        private string _cachedToken;

        private void LoadCachedTokenFromFile()
        {
            try
            {
                if (File.Exists("Token-v9.txt"))
                {
                    var token = File.ReadAllText("Token-v9.txt").Trim();
                    if (!string.IsNullOrEmpty(token))
                    {
                        _cachedToken = token;
                        btnMe.Enabled = btnUsers.Enabled = btnApplications.Enabled = btnGroups.Enabled = true;
                        btnSignIn.Text = "Signed In";
                        
                        AppendLine("Loaded cached token from file.");
                    }
                }
            }
            catch (Exception ex)
            {
                AppendLine($"Failed to load cached token: {ex.Message}");
            }
        }

        private async void BtnMe_Click(object sender, EventArgs e)
        {
            txtOutput.Clear();
            if (string.IsNullOrEmpty(_cachedToken))
            {
                AppendLine("Please sign in first using the Sign In button.");
                return;
            }

            await FetchAndDisplay("/me", "https://graph.microsoft.com/v1.0/me");
        }

        private async void BtnUsers_Click(object sender, EventArgs e)
        {
            txtOutput.Clear();
            if (string.IsNullOrEmpty(_cachedToken))
            {
                AppendLine("Please sign in first using the Sign In button.");
                return;
            }

            await FetchAndPopulateUsersAsync("https://graph.microsoft.com/v1.0/users");
        }

        private async void BtnApplications_Click(object sender, EventArgs e)
        {
            txtOutput.Clear();
            if (string.IsNullOrEmpty(_cachedToken))
            {
                AppendLine("Please sign in first using the Sign In button.");
                return;
            }

            await FetchAndDisplay("/applications", "https://graph.microsoft.com/v1.0/applications");
        }

        private async void BtnGroups_Click(object sender, EventArgs e)
        {
            txtOutput.Clear();
            if (string.IsNullOrEmpty(_cachedToken))
            {
                AppendLine("Please sign in first using the Sign In button.");
                return;
            }

            await FetchAndDisplay("/groups", "https://graph.microsoft.com/v1.0/groups");
        }

        // Explicit Sign In button to acquire token without calling an endpoint
        private async void BtnSignIn_Click(object sender, EventArgs e)
        {
            txtOutput.Clear();

            await EnsureTokenAsync();

            if (!string.IsNullOrEmpty(_cachedToken))
            {
                btnMe.Enabled = btnUsers.Enabled = btnApplications.Enabled = btnGroups.Enabled = true;
                btnSignIn.Text = "Signed In";
               
            }
        }

        private async Task FetchAndDisplay(string header, string url)
        {
            try
            {
                using (var http = new HttpClient())
                {
                    http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _cachedToken);
                    AppendHeader(header);
                    var json = await GetGraphAsync(http, url);
                    AppendLine(json);
                }
            }
            catch (HttpRequestException hex)
            {
                AppendLine($"HTTP error: {hex.Message}");
            }
            catch (Exception ex)
            {
                AppendLine($"Unexpected error: {ex.Message}");
            }
        }

        // NEW: fetch users, populate the right-hand listbox
        private async Task FetchAndPopulateUsersAsync(string url)
        {
            try
            {
                using (var http = new HttpClient())
                {
                    http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _cachedToken);
                    AppendHeader("/users");
                    var json = await GetGraphAsync(http, url);

                    // Parse and populate list
                    lstUsers.Items.Clear();

                    using (var doc = JsonDocument.Parse(json))
                    {
                        if (doc.RootElement.TryGetProperty("value", out var array) && array.ValueKind == JsonValueKind.Array)
                        {
                            var items = new List<UserItem>();
                            foreach (var el in array.EnumerateArray())
                            {
                                string id = el.TryGetProperty("id", out var pid) ? pid.GetString() : null;
                                string display = el.TryGetProperty("displayName", out var pd) ? pd.GetString() : null;
                                string upn = el.TryGetProperty("userPrincipalName", out var pupn) ? pupn.GetString() : null;

                                items.Add(new UserItem { Id = id, Display = display, UserPrincipalName = upn });
                            }

                            // Add to listbox
                            foreach (var ui in items)
                                lstUsers.Items.Add(ui);

                            AppendLine($"Loaded {items.Count} users. Select one on the right to fetch details.");
                        }
                        else
                        {
                            // fallback: show raw
                            AppendLine(json);
                        }
                    }
                }
            }
            catch (HttpRequestException hex)
            {
                AppendLine($"HTTP error: {hex.Message}");
            }
            catch (Exception ex)
            {
                AppendLine($"Unexpected error: {ex.Message}");
            }
        }

        // When a user is selected, fetch that user's detail
        private async void LstUsers_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstUsers.SelectedItem is UserItem item && !string.IsNullOrEmpty(item.Id))
            {
                string url = $"https://graph.microsoft.com/beta/users/{Uri.EscapeDataString(item.Id)}";
                AppendHeader($"/users/{item.Id}");
                try
                {
                    using (var http = new HttpClient())
                    {
                        http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _cachedToken);
                        var json = await GetGraphAsync(http, url);
                        AppendLine(json);
                    }
                }
                catch (Exception ex)
                {
                    AppendLine($"Failed to fetch user: {ex.Message}");
                }
            }
        }

        private void AppendHeader(string title)
        {
            AppendLine("--------------- " + title + " ---------------");
        }

        private void AppendLine(string text)
        {
            if (txtOutput.InvokeRequired)
            {
                txtOutput.Invoke(new Action(() => { txtOutput.AppendText(text + Environment.NewLine); }));
            }
            else
            {
                txtOutput.AppendText(text + Environment.NewLine);
            }
        }

        private async Task<string> GetGraphAsync(HttpClient http, string url)
        {
            var response = await http.GetAsync(url);
            response.EnsureSuccessStatusCode();

            using (var stream = await response.Content.ReadAsStreamAsync())
            using (var doc = await JsonDocument.ParseAsync(stream))
            {
                return JsonSerializer.Serialize(doc, new JsonSerializerOptions { WriteIndented = true });
            }
        }

        private void Form1_Load_1(object sender, EventArgs e)
        {

        }

        // small helper to keep objects in the ListBox with a friendly string
        private class UserItem
        {
            public string Id { get; set; }
            public string Display { get; set; }
            public string UserPrincipalName { get; set; }
            public override string ToString()
            {
                if (!string.IsNullOrEmpty(Display)) return Display;
                if (!string.IsNullOrEmpty(UserPrincipalName)) return UserPrincipalName;
                return Id ?? "(unknown)";
            }
        }
    }
}