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
            MicrosoftGraphService.Instance.AuthenticationModel = MicrosoftGraphEnums.AuthenticationModel.V2;
            MicrosoftGraphService.Instance.SignInFailed += async (ss, se) =>
            {
                var error = new MessageDialog(se.Exception.ToString());
                await error.ShowAsync();
            };

            // Initialize the service.  We really only need read access.
            string[] scopes = { "Notes.Read" };
            if (!MicrosoftGraphService.Instance.Initialize(
                    clientId,
                    MicrosoftGraphEnums.ServicesToInitialize.Message | MicrosoftGraphEnums.ServicesToInitialize.UserProfile | MicrosoftGraphEnums.ServicesToInitialize.Event,
                    scopes))
            {
                return;
            }

            // Login
            try
            {
                if (!await MicrosoftGraphService.Instance.LoginAsync())
                {
                    var error = new MessageDialog("Unable to sign in");
                    await error.ShowAsync();
                    return;
                }
            }
            catch (AdalServiceException ase)
            {
                var error = new MessageDialog(ase.Message);
                await error.ShowAsync();
                return;
            }
            catch (AdalException ae)
            {
                var error = new MessageDialog(ae.Message);
                await error.ShowAsync();
                return;
            }
            catch (MsalServiceException mse)
            {
                var error = new MessageDialog(mse.Message);
                await error.ShowAsync();
                return;
            }
            catch (MsalException me)
            {
                var error = new MessageDialog(me.Message);
                await error.ShowAsync();
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

                    this.Preview.NavigateToString(contentStr);
                }
            }
        }
    }
}
