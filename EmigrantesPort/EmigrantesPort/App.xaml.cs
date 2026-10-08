namespace EmigrantesPort;

public partial class App : Application
{
    private readonly IWorkbookDestination destination;
    public App(IWorkbookDestination destination)
    {
        InitializeComponent(); this.destination = destination;
        UserAppTheme = AppTheme.Light;
    }
    protected override Window CreateWindow(IActivationState? activationState) => new(new NavigationPage(new MainPage(destination))
    { BarBackgroundColor = Color.FromArgb("#F7F8F4"), BarTextColor = Color.FromArgb("#173F35") })
    { Title = "Emigrantes · Portugal", Width = 1380, Height = 900, MinimumWidth = 1050, MinimumHeight = 720 };
}
