
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
        public MainPage()
        {
            this.InitializeComponent();

            this.NotePicker.OnBusyStart += (sender, args) => this.BusyIndicator.IsActive = true;
            this.NotePicker.OnBusyEnd += (sender, args) => this.BusyIndicator.IsActive = false;
        }

        private async void LoginButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(this.ClientId.Text))
            {
                await this.NotePicker.Start(this.ClientId.Text);
            }
        }
    }
}
