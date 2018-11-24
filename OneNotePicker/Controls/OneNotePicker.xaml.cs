// --------------------------------------------------------------------------------------------------------------------
// <copyright company="otherslikeyou.com Inc." file="OneNotePicker.xaml.cs">
//   Licensed under the MIT License. See LICENSE file in the project root for full license information.
// </copyright>
// <summary>
//   
// </summary>
// 
// --------------------------------------------------------------------------------------------------------------------

namespace OneNotePicker.Controls
{
    using System;
    using System.IO;
    using System.Threading.Tasks;
    using Microsoft.Graph;
    using Microsoft.Identity.Client;
    using Microsoft.IdentityModel.Clients.ActiveDirectory;
    using Microsoft.Toolkit.Services.MicrosoftGraph;
    using Windows.UI.Popups;
    using Windows.UI.Xaml.Controls;

    /// <summary>
    /// The one note picker.
    /// </summary>
    public sealed partial class OneNotePicker
    {
        public EventHandler OnLoggedIn;

        public EventHandler OnLoggedOut;

        public EventHandler OnBusyStart;

        public EventHandler OnBusyEnd;

        /// <summary>
        /// Initializes a new instance of the <see cref="OneNotePicker"/> class.
        /// </summary>
        public OneNotePicker()
        {
            this.InitializeComponent();

            this.NotebooksList.SelectionChanged += this.NotebooksListOnSelectionChanged;
            this.SectionsList.SelectionChanged += this.SectionsListOnSelectionChanged;
            this.PagesList.SelectionChanged += this.PagesListOnSelectionChanged;
        }

        /// <summary>
        /// Initialize the MSGraph session, populate the notebooks list box.
        /// </summary>
        /// <param name="clientId">The App ID / Client Id.</param>
        /// <returns>
        /// The <see cref="Task"/>.
        /// </returns>
        public async Task Start(string clientId)
        {
            bool loginSuccess = false;

            MicrosoftGraphService.Instance.AuthenticationModel = MicrosoftGraphEnums.AuthenticationModel.V2;
            MicrosoftGraphService.Instance.SignInFailed += async (ss, se) =>
            {
                var error = new MessageDialog(se.Exception.ToString());
                await error.ShowAsync();
            };

            // Initialize the service.  We really only need read access.
            string[] scopes = { "User.Read", "Notes.Read" };
            if (!MicrosoftGraphService.Instance.Initialize(
                    clientId,
                    MicrosoftGraphEnums.ServicesToInitialize.Message | MicrosoftGraphEnums.ServicesToInitialize.UserProfile | MicrosoftGraphEnums.ServicesToInitialize.Event,
                    scopes))
            {
                return;
            }

            this.OnBusyStart?.Invoke(this, EventArgs.Empty);

            // Login
            try
            {
                if (!await MicrosoftGraphService.Instance.LoginAsync())
                {
                    var error = new MessageDialog("Unable to sign in");
                    await error.ShowAsync();
                    return;
                }

                loginSuccess = true;
            }
            catch (AdalServiceException ase)
            {
                var error = new MessageDialog(ase.Message);
                await error.ShowAsync();
            }
            catch (AdalException ae)
            {
                var error = new MessageDialog(ae.Message);
                await error.ShowAsync();
            }
            catch (MsalServiceException mse)
            {
                var error = new MessageDialog(mse.Message);
                await error.ShowAsync();
            }
            catch (MsalException me)
            {
                var error = new MessageDialog(me.Message);
                await error.ShowAsync();
            }
            catch (Exception ex)
            {
                var error = new MessageDialog(ex.Message);
                await error.ShowAsync();
            }

            if (!loginSuccess)
            {
                // Just in case it was a perms error, let's not leave the user logged in
                try
                {
                    await MicrosoftGraphService.Instance.Logout();
                }
                catch (Exception)
                {
                    // ignored
                }

                this.OnBusyEnd?.Invoke(this, EventArgs.Empty);
                return;
            }

            // Populate the list of notebooks
            var notebooks = await MicrosoftGraphService
                .Instance
                .GraphProvider
                .Me
                .Onenote
                .Notebooks
                .Request()
                .GetAsync();

            if (this.NotebooksList.Items != null)
            {
                foreach (var notebook in notebooks)
                {
                    this.NotebooksList.Items.Add(notebook);
                }
            }

            this.OnBusyEnd?.Invoke(this, EventArgs.Empty);
            this.OnLoggedIn?.Invoke(this, EventArgs.Empty);
        }

        /// <summary>
        /// A notebook was selected from the ListView.  Populate the sections for it.
        /// </summary>
        /// <param name="sender">The event sender.</param>
        /// <param name="e">The event arguments.</param>
        private async void NotebooksListOnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count > 0)
            {
                if (e.AddedItems[0] is Notebook notebook)
                {
                    this.OnBusyStart?.Invoke(this, EventArgs.Empty);

                    var sections = await MicrosoftGraphService
                                       .Instance
                                       .GraphProvider
                                       .Me
                                       .Onenote
                                       .Notebooks[notebook.Id]
                                       .Sections
                                       .Request()
                                       .GetAsync();

                    this.PagesList.Items?.Clear();

                    if (this.SectionsList.Items != null)
                    {
                        this.SectionsList.Items.Clear();
                        foreach (var section in sections)
                        {
                            this.SectionsList.Items.Add(section);
                        }
                    }

                    this.OnBusyEnd?.Invoke(this, EventArgs.Empty);
                }
            }
        }

        /// <summary>
        /// A section was selected from the ListView.  Populate the pages for it.
        /// </summary>
        /// <param name="sender">The event sender.</param>
        /// <param name="e">The event arguments.</param>
        private async void SectionsListOnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count > 0)
            {
                if (e.AddedItems[0] is OnenoteSection section)
                {
                    this.OnBusyStart?.Invoke(this, EventArgs.Empty);

                    var pages = await MicrosoftGraphService
                                    .Instance
                                    .GraphProvider
                                    .Me
                                    .Onenote
                                    .Sections[section.Id]
                                    .Pages
                                    .Request()
                                    .GetAsync();

                    if (this.PagesList.Items != null)
                    {
                        this.PagesList.Items.Clear();

                        foreach (var page in pages)
                        {
                            this.PagesList.Items.Add(page);
                        }
                    }

                    this.OnBusyEnd?.Invoke(this, EventArgs.Empty);
                }
            }
        }

        /// <summary>
        /// A page was selected from the ListView.  Retrieve the content and display it in the WebView.
        /// </summary>
        /// <param name="sender">The event sender.</param>
        /// <param name="e">The event arguments.</param>
        private async void PagesListOnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count > 0)
            {
                if (e.AddedItems[0] is OnenotePage page)
                {
                    this.OnBusyStart?.Invoke(this, EventArgs.Empty);

                    var contentStream = await MicrosoftGraphService
                                            .Instance
                                            .GraphProvider
                                            .Me
                                            .Onenote
                                            .Pages[page.Id]
                                            .Content
                                            .Request()
                                            .GetAsync();

                    string contentStr;

                    using (var reader = new StreamReader(contentStream))
                    {
                        contentStr = reader.ReadToEnd();
                    }

                    this.PageTitle.Text = page.Title;
                    if (page.LastModifiedDateTime != null)
                    {
                        this.PageDate.Text = ((DateTimeOffset)page.LastModifiedDateTime).DateTime.ToLongDateString() +
                                             " " +
                                             ((DateTimeOffset)page.LastModifiedDateTime).DateTime.ToShortTimeString();
                    }

                    this.Preview.NavigateToString(contentStr);

                    this.OnBusyEnd?.Invoke(this, EventArgs.Empty);
                }
            }
        }
    }
}
