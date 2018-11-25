
// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace OneNotePicker
{
    using Windows.UI.Xaml;
    using Windows.UI.Xaml.Controls;

    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        private bool loggedIn;

        public MainPage()
        {
            this.InitializeComponent();

            this.NotePicker.OnBusyStart += (sender, args) => this.BusyIndicator.IsActive = true;
            this.NotePicker.OnBusyEnd += (sender, args) => this.BusyIndicator.IsActive = false;
            this.NotePicker.OnLoggedIn += (sender, args) =>
                {
                    this.loggedIn = true;
                    this.LoginLogoutButton.Content = "Log Out";
                };
            this.NotePicker.OnLoggedOut += (sender, args) =>
                {
                    this.loggedIn = false;
                    this.LoginLogoutButton.Content = "Log In";
                };
        }

        private async void LoginLogoutButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (!this.loggedIn)
            {
                if (!string.IsNullOrEmpty(this.ClientId.Text))
                {
                    await this.NotePicker.Start(this.ClientId.Text);
                }
            }
            else
            {
                this.NotePicker.Logout();
            }
        }
    }
}
