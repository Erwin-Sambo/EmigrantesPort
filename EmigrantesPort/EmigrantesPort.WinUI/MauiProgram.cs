namespace EmigrantesPort.WinUI
{
    public static class MauiProgram
    {
        public static MauiApp CreateMauiApp()
        {
            var builder = MauiApp.CreateBuilder();
            builder.Services.AddSingleton<IWorkbookDestination, WindowsWorkbookDestination>();

            builder
                .UseSharedMauiApp();

            return builder.Build();
        }
    }
}
