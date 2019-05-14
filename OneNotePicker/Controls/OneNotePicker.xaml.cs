// --------------------------------------------------------------------------------------------------------------------
// <copyright company="otherslikeyou.com Inc." file="OneNotePicker.xaml.cs">
//   Licensed under the MIT License. See LICENSE file in the project root for full license information.
// </copyright>
// <summary>
//   
// </summary>
// 
// --------------------------------------------------------------------------------------------------------------------

namespace OLY.OneNotePicker.Controls
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Threading.Tasks;
    using Microsoft.Graph;
    using Microsoft.Identity.Client;
    using Microsoft.IdentityModel.Clients.ActiveDirectory;
    using Microsoft.Toolkit.Services.MicrosoftGraph;
    using OLY.OneNotePicker.Models;
    using Windows.UI.Popups;
    using Windows.UI.Xaml.Controls;

    /// <summary>
    /// The one note picker.
    /// </summary>
    public sealed partial class OneNotePicker
    {
        /// <summary>
        /// Invoked when the user has been successfully logged in.
        /// </summary>
        public EventHandler OnLoggedIn;

        /// <summary>
        /// Invoked when the user has been logged out.
        /// </summary>
        public EventHandler OnLoggedOut;

        /// <summary>
        /// Invoked when a Graph operation begins.
        /// </summary>
        public EventHandler OnBusyStart;

        /// <summary>
        /// Invoked when a Graph operation ends.
        /// </summary>
        public EventHandler OnBusyEnd;

        /// <summary>
        /// Invoked when the selected OneNote page changes.
        /// </summary>
        public EventHandler SelectedPageChanged;

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

        public SelectedPage SelectedPage { get; private set; }

        /// <summary>
        /// Initialize the MSGraph session, populate the notebooks list box.
        /// </summary>
        /// <param name="clientId">The App ID / Client Id.</param>
        /// <returns>
        /// The <see cref="Task"/>.
        /// </returns>
        public async Task<bool> Start(string clientId)
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
                return false;
            }

            this.OnBusyStart?.Invoke(this, EventArgs.Empty);

            // Login
            try
            {
                if (!await MicrosoftGraphService.Instance.LoginAsync())
                {
                    var error = new MessageDialog("Unable to sign in");
                    await error.ShowAsync();
                    return false;
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
                this.Logout();
                this.OnBusyEnd?.Invoke(this, EventArgs.Empty);
                return false;
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

            return true;
        }

        /// <summary>
        /// Log the user out.
        /// </summary>
        public async void Logout()
        {
            try
            {
                await MicrosoftGraphService.Instance.Logout();
                this.OnLoggedOut?.Invoke(this, EventArgs.Empty);
            }
            catch (Exception)
            {
                // ignored
            }
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

                    List<QueryOption> contentOptions =
                        new List<QueryOption> { new QueryOption("preAuthenticated", "true") };

                    var contentStream = await MicrosoftGraphService
                                            .Instance
                                            .GraphProvider
                                            .Me
                                            .Onenote
                                            .Pages[page.Id]
                                            .Content
                                            .Request(contentOptions)
                                            .GetAsync();

                    string contentStr;

                    using (var reader = new StreamReader(contentStream))
                    {
                        contentStr = reader.ReadToEnd();
                    }

                    this.SelectedPage = new SelectedPage
                                            {
                                                Title = page.Title,
                                                Modified = page.LastModifiedDateTime?.DateTime ?? DateTime.MinValue,
                                                Content = contentStr
                                            };

                    this.PageTitle.Text = this.SelectedPage.Title;
                    if (this.SelectedPage.Modified != DateTime.MinValue)
                    {
                        this.PageDate.Text = this.SelectedPage.Modified.ToLongDateString() +
                                             " " +
                                             this.SelectedPage.Modified.ToShortTimeString();
                    }
                    else
                    {
                        this.PageDate.Text = string.Empty;
                    }

                    this.Preview.NavigateToString(this.SelectedPage.Content);

                    this.OnBusyEnd?.Invoke(this, EventArgs.Empty);
                    this.SelectedPageChanged?.Invoke(this,EventArgs.Empty);
                }
            }
        }
    }
}
