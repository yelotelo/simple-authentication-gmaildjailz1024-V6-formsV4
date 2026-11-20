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
        private Button btnServicePrincipals;
        private Button btnGroups;
        private Button btnDirectoryRoles;
        private Button btnToken;
        private Button btnSignIn;
        private ComboBox cmbClientIds;
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
            btnServicePrincipals = new Button { Text = "Service Principals", Width = 140, Height = 34, Margin = new Padding(6) };
            btnGroups = new Button { Text = "Groups", Width = 120, Height = 34, Margin = new Padding(6) };
            btnDirectoryRoles = new Button { Text = "DirectoryRoles", Width = 130, Height = 34, Margin = new Padding(6) };
            btnToken = new Button { Text = "Token", Width = 120, Height = 34, Margin = new Padding(6) };
            cmbClientIds = new ComboBox { Width = 280, Height = 34, Margin = new Padding(6), DropDownStyle = ComboBoxStyle.DropDownList };

            cmbClientIds.Items.Add(new ClientIdOption { Name = "Microsoft Graph Command Line Tools", ClientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e" });
            cmbClientIds.Items.Add(new ClientIdOption { Name = "Azure Active Directory PowerShell", ClientId = "1b730954-1685-4b74-9bfd-dac224a7b894" });
            cmbClientIds.Items.Add(new ClientIdOption { Name = "Microsoft Office", ClientId = "d3590ed6-52b3-4102-aeff-aad2292ab01c" });
            cmbClientIds.Items.Add(new ClientIdOption { Name = "Microsoft Azure CLI", ClientId = "04b07795-8ddb-461a-bbee-02f9e1bf7b46" });
            cmbClientIds.Items.Add(new ClientIdOption { Name = "Microsoft Teams", ClientId = "1fec8e78-bce4-4aaf-ab1b-5451cc387264" });
            cmbClientIds.Items.Add(new ClientIdOption { Name = "Microsoft Azure PowerShell", ClientId = "1950a258-227b-4e31-a9cf-717495945fc2" });
            cmbClientIds.Items.Add(new ClientIdOption { Name = "Microsoft Edge", ClientId = "ecd6b820-32c2-49b6-98a6-444530e5a77a" });
            cmbClientIds.Items.Add(new ClientIdOption { Name = "Visual Studio Code", ClientId = "aebc6443-996d-45c2-90f0-388ff96faa56" });
            cmbClientIds.SelectedIndex = 0;

            panel.Controls.Add(btnMe);
            panel.Controls.Add(btnUsers);
            panel.Controls.Add(btnApplications);
            panel.Controls.Add(btnServicePrincipals);
            panel.Controls.Add(btnGroups);
            panel.Controls.Add(btnDirectoryRoles);
            panel.Controls.Add(btnToken);
            panel.Controls.Add(cmbClientIds);
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
            btnServicePrincipals.Click += BtnServicePrincipals_Click;
            btnGroups.Click += BtnGroups_Click;
            btnDirectoryRoles.Click += BtnDirectoryRoles_Click;
            btnToken.Click += BtnToken_Click;
            btnSignIn.Click += BtnSignIn_Click;

            // Start with action buttons disabled until explicit sign-in
            btnMe.Enabled = btnUsers.Enabled = btnApplications.Enabled = btnServicePrincipals.Enabled = btnGroups.Enabled = btnDirectoryRoles.Enabled = btnToken.Enabled = false;

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
            btnMe.Enabled = btnUsers.Enabled = btnApplications.Enabled = btnServicePrincipals.Enabled = btnGroups.Enabled = btnDirectoryRoles.Enabled = btnToken.Enabled = false;
            cmbClientIds.Enabled = false;
            txtOutput.Clear();

            try
            {
                string tenant = "djailz1024gmail.onmicrosoft.com";
                var selectedClient = cmbClientIds.SelectedItem as ClientIdOption;
                if (selectedClient == null)
                {
                    AppendLine("Please choose a client app before signing in.");
                    return;
                }
                string clientId = selectedClient.ClientId;

                string authority = $"https://login.microsoftonline.com/{tenant}";

                string[] scopes = new[]
                {
                ""

                };


                /*
                string[] scopes = new[]
                {
                    "offline_access",
                    "group.read.all",
                    "openid",
                    "User.Read.All"
                }; */

                var app = PublicClientApplicationBuilder
                    .Create(clientId)
                    .WithAuthority(authority)
                    .WithRedirectUri("https://login.microsoftonline.com/common/oauth2/nativeclient")
                    .Build();

                AppendLine($"Using client: {selectedClient.Name} ({clientId})");
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
                btnMe.Enabled = btnUsers.Enabled = btnApplications.Enabled = btnServicePrincipals.Enabled = btnGroups.Enabled = btnDirectoryRoles.Enabled = btnToken.Enabled = true;
                cmbClientIds.Enabled = true;
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
                        btnMe.Enabled = btnUsers.Enabled = btnApplications.Enabled = btnServicePrincipals.Enabled = btnGroups.Enabled = btnDirectoryRoles.Enabled = btnToken.Enabled = true;
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

            await FetchAndPopulateApplicationsAsync("https://graph.microsoft.com/v1.0/applications");
        }

        private async void BtnGroups_Click(object sender, EventArgs e)
        {
            txtOutput.Clear();
            if (string.IsNullOrEmpty(_cachedToken))
            {
                AppendLine("Please sign in first using the Sign In button.");
                return;
            }

            await FetchAndPopulateGroupsAsync("https://graph.microsoft.com/v1.0/groups");
        }

        private async void BtnServicePrincipals_Click(object sender, EventArgs e)
        {
            txtOutput.Clear();
            if (string.IsNullOrEmpty(_cachedToken))
            {
                AppendLine("Please sign in first using the Sign In button.");
                return;
            }

            await FetchAndPopulateServicePrincipalsAsync("https://graph.microsoft.com/v1.0/servicePrincipals");
        }

        private async void BtnDirectoryRoles_Click(object sender, EventArgs e)
        {
            txtOutput.Clear();
            if (string.IsNullOrEmpty(_cachedToken))
            {
                AppendLine("Please sign in first using the Sign In button.");
                return;
            }

            await FetchAndPopulateDirectoryRolesAsync("https://graph.microsoft.com/v1.0/directoryRoles");
        }

        private void BtnToken_Click(object sender, EventArgs e)
        {
            txtOutput.Clear();
            if (string.IsNullOrEmpty(_cachedToken))
            {
                AppendLine("Please sign in first using the Sign In button.");
                return;
            }

            DecodeAndDisplayToken(_cachedToken);
        }

        // Explicit Sign In button to acquire token without calling an endpoint
        private async void BtnSignIn_Click(object sender, EventArgs e)
        {
            txtOutput.Clear();

            await EnsureTokenAsync();

            if (!string.IsNullOrEmpty(_cachedToken))
            {
                btnMe.Enabled = btnUsers.Enabled = btnApplications.Enabled = btnServicePrincipals.Enabled = btnGroups.Enabled = btnDirectoryRoles.Enabled = btnToken.Enabled = true;
                btnSignIn.Text = "Signed In";
               
            }
        }

        private void DecodeAndDisplayToken(string token)
        {
            try
            {
                var parts = token.Split('.');
                if (parts.Length < 2)
                {
                    AppendLine("Invalid JWT format.");
                    return;
                }

                var headerJson = DecodeBase64Url(parts[0]);
                var payloadJson = DecodeBase64Url(parts[1]);

                AppendHeader("Encrypted TOKEN");
                AppendLine(token);

                AppendHeader("TOKEN HEADER");
                AppendLine(PrettyJson(headerJson));

                AppendHeader("TOKEN PAYLOAD");
                AppendLine(PrettyJson(payloadJson));

                using (var doc = JsonDocument.Parse(payloadJson))
                {
                    var root = doc.RootElement;
                    if (root.TryGetProperty("iat", out var iat) && iat.ValueKind == JsonValueKind.Number)
                    {
                        AppendLine($"iat: {UnixToUtc(iat.GetInt64()):u}");
                    }
                    if (root.TryGetProperty("nbf", out var nbf) && nbf.ValueKind == JsonValueKind.Number)
                    {
                        AppendLine($"nbf: {UnixToUtc(nbf.GetInt64()):u}");
                    }
                    if (root.TryGetProperty("exp", out var exp) && exp.ValueKind == JsonValueKind.Number)
                    {
                        AppendLine($"exp: {UnixToUtc(exp.GetInt64()):u}");
                    }
                }
            }
            catch (Exception ex)
            {
                AppendLine($"Failed to decode token: {ex.Message}");
            }
        }

        private string DecodeBase64Url(string input)
        {
            var s = input.Replace('-', '+').Replace('_', '/');
            switch (s.Length % 4)
            {
                case 2: s += "=="; break;
                case 3: s += "="; break;
            }
            var bytes = Convert.FromBase64String(s);
            return Encoding.UTF8.GetString(bytes);
        }

        private string PrettyJson(string json)
        {
            using (var doc = JsonDocument.Parse(json))
            {
                return JsonSerializer.Serialize(doc, new JsonSerializerOptions { WriteIndented = true });
            }
        }

        private DateTime UnixToUtc(long seconds)
        {
            return DateTimeOffset.FromUnixTimeSeconds(seconds).UtcDateTime;
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

        // NEW: fetch applications, populate the right-hand listbox
        private async Task FetchAndPopulateApplicationsAsync(string url)
        {
            try
            {
                using (var http = new HttpClient())
                {
                    http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _cachedToken);
                    AppendHeader("/applications");
                    var json = await GetGraphAsync(http, url);

                    // Parse and populate list
                    lstUsers.Items.Clear();

                    using (var doc = JsonDocument.Parse(json))
                    {
                        if (doc.RootElement.TryGetProperty("value", out var array) && array.ValueKind == JsonValueKind.Array)
                        {
                            var items = new List<ApplicationItem>();
                            foreach (var el in array.EnumerateArray())
                            {
                                string id = el.TryGetProperty("id", out var pid) ? pid.GetString() : null;
                                string display = el.TryGetProperty("displayName", out var pd) ? pd.GetString() : null;
                                string appId = el.TryGetProperty("appId", out var pappId) ? pappId.GetString() : null;

                                items.Add(new ApplicationItem { Id = id, Display = display, AppId = appId });
                            }

                            // Add to listbox
                            foreach (var ai in items)
                                lstUsers.Items.Add(ai);

                            AppendLine($"Loaded {items.Count} applications. Select one on the right to fetch details.");
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

        // NEW: fetch groups, populate the right-hand listbox
        private async Task FetchAndPopulateGroupsAsync(string url)
        {
            try
            {
                using (var http = new HttpClient())
                {
                    http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _cachedToken);
                    AppendHeader("/groups");
                    var json = await GetGraphAsync(http, url);

                    // Parse and populate list
                    lstUsers.Items.Clear();

                    using (var doc = JsonDocument.Parse(json))
                    {
                        if (doc.RootElement.TryGetProperty("value", out var array) && array.ValueKind == JsonValueKind.Array)
                        {
                            var items = new List<GroupItem>();
                            foreach (var el in array.EnumerateArray())
                            {
                                string id = el.TryGetProperty("id", out var pid) ? pid.GetString() : null;
                                string display = el.TryGetProperty("displayName", out var pd) ? pd.GetString() : null;

                                items.Add(new GroupItem { Id = id, Display = display });
                            }

                            // Add to listbox
                            foreach (var gi in items)
                                lstUsers.Items.Add(gi);

                            AppendLine($"Loaded {items.Count} groups. Select one on the right to fetch details.");
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

        // NEW: fetch service principals, populate the right-hand listbox
        private async Task FetchAndPopulateServicePrincipalsAsync(string url)
        {
            try
            {
                using (var http = new HttpClient())
                {
                    http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _cachedToken);
                    AppendHeader("/servicePrincipals");
                    var json = await GetGraphAsync(http, url);

                    // Parse and populate list
                    lstUsers.Items.Clear();

                    using (var doc = JsonDocument.Parse(json))
                    {
                        if (doc.RootElement.TryGetProperty("value", out var array) && array.ValueKind == JsonValueKind.Array)
                        {
                            var items = new List<ServicePrincipalItem>();
                            foreach (var el in array.EnumerateArray())
                            {
                                string id = el.TryGetProperty("id", out var pid) ? pid.GetString() : null;
                                string display = el.TryGetProperty("displayName", out var pd) ? pd.GetString() : null;
                                string appId = el.TryGetProperty("appId", out var pappId) ? pappId.GetString() : null;

                                items.Add(new ServicePrincipalItem { Id = id, Display = display, AppId = appId });
                            }

                            foreach (var si in items)
                                lstUsers.Items.Add(si);

                            AppendLine($"Loaded {items.Count} service principals. Select one on the right to fetch details.");
                        }
                        else
                        {
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

        // NEW: fetch directory roles, populate the right-hand listbox
        private async Task FetchAndPopulateDirectoryRolesAsync(string url)
        {
            try
            {
                using (var http = new HttpClient())
                {
                    http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _cachedToken);
                    AppendHeader("/directoryRoles");
                    var json = await GetGraphAsync(http, url);

                    // Parse and populate list
                    lstUsers.Items.Clear();

                    using (var doc = JsonDocument.Parse(json))
                    {
                        if (doc.RootElement.TryGetProperty("value", out var array) && array.ValueKind == JsonValueKind.Array)
                        {
                            var items = new List<DirectoryRoleItem>();
                            foreach (var el in array.EnumerateArray())
                            {
                                string id = el.TryGetProperty("id", out var pid) ? pid.GetString() : null;
                                string display = el.TryGetProperty("displayName", out var pd) ? pd.GetString() : null;
                                string roleTemplateId = el.TryGetProperty("roleTemplateId", out var prt) ? prt.GetString() : null;

                                items.Add(new DirectoryRoleItem { Id = id, Display = display, RoleTemplateId = roleTemplateId });
                            }

                            foreach (var dr in items)
                                lstUsers.Items.Add(dr);

                            AppendLine($"Loaded {items.Count} directory roles. Select one on the right to fetch details.");
                        }
                        else
                        {
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

        // When a user, application, group, service principal, or directory role is selected, fetch that item's detail
        private async void LstUsers_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstUsers.SelectedItem is UserItem userItem && !string.IsNullOrEmpty(userItem.Id))
            {
                string url = $"https://graph.microsoft.com/beta/users/{Uri.EscapeDataString(userItem.Id)}";
                AppendHeader($"/users/{userItem.Id}");

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
            else if (lstUsers.SelectedItem is ApplicationItem appItem && !string.IsNullOrEmpty(appItem.Id))
            {
                string url = $"https://graph.microsoft.com/v1.0/applications/{Uri.EscapeDataString(appItem.Id)}";
                AppendHeader($"/applications/{appItem.Id}");

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
                    AppendLine($"Failed to fetch application: {ex.Message}");
                }
            }
            else if (lstUsers.SelectedItem is GroupItem groupItem && !string.IsNullOrEmpty(groupItem.Id))
            {
                string url = $"https://graph.microsoft.com/v1.0/groups/{Uri.EscapeDataString(groupItem.Id)}";
                AppendHeader($"/groups/{groupItem.Id}");

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
                    AppendLine($"Failed to fetch group: {ex.Message}");
                }
            }
            else if (lstUsers.SelectedItem is ServicePrincipalItem servicePrincipalItem && !string.IsNullOrEmpty(servicePrincipalItem.Id))
            {
                string url = $"https://graph.microsoft.com/v1.0/servicePrincipals/{Uri.EscapeDataString(servicePrincipalItem.Id)}";
                AppendHeader($"/servicePrincipals/{servicePrincipalItem.Id}");

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
                    AppendLine($"Failed to fetch service principal: {ex.Message}");
                }
            }
            else if (lstUsers.SelectedItem is DirectoryRoleItem directoryRoleItem && !string.IsNullOrEmpty(directoryRoleItem.Id))
            {
                string url = $"https://graph.microsoft.com/v1.0/directoryRoles/{Uri.EscapeDataString(directoryRoleItem.Id)}";
                AppendHeader($"/directoryRoles/{directoryRoleItem.Id}");

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
                    AppendLine($"Failed to fetch directory role: {ex.Message}");
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

        // Helper class for applications
        private class ApplicationItem
        {
            public string Id { get; set; }
            public string Display { get; set; }
            public string AppId { get; set; }
            public override string ToString()
            {
                if (!string.IsNullOrEmpty(Display)) return Display;
                if (!string.IsNullOrEmpty(AppId)) return AppId;
                return Id ?? "(unknown)";
            }
        }

        // Helper class for groups
        private class GroupItem
        {
            public string Id { get; set; }
            public string Display { get; set; }
            public override string ToString()
            {
                if (!string.IsNullOrEmpty(Display)) return Display;
                return Id ?? "(unknown)";
            }
        }

        // Helper class for service principals
        private class ServicePrincipalItem
        {
            public string Id { get; set; }
            public string Display { get; set; }
            public string AppId { get; set; }
            public override string ToString()
            {
                if (!string.IsNullOrEmpty(Display)) return Display;
                if (!string.IsNullOrEmpty(AppId)) return AppId;
                return Id ?? "(unknown)";
            }
        }

        // Helper class for directory roles
        private class DirectoryRoleItem
        {
            public string Id { get; set; }
            public string Display { get; set; }
            public string RoleTemplateId { get; set; }
            public override string ToString()
            {
                if (!string.IsNullOrEmpty(Display)) return Display;
                if (!string.IsNullOrEmpty(RoleTemplateId)) return RoleTemplateId;
                return Id ?? "(unknown)";
            }
        }

        private class ClientIdOption
        {
            public string Name { get; set; }
            public string ClientId { get; set; }
            public override string ToString()
            {
                return $"{Name} ({ClientId})";
            }
        }
    }
}