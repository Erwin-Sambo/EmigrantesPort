using Windows.Storage.Pickers;

namespace EmigrantesPort.WinUI;

public sealed class WindowsWorkbookDestination : IWorkbookDestination
{
    public async Task<string?> SaveCopyAsync(byte[] workbook, string suggestedName, string currentPath)
    {
        var window = Microsoft.Maui.Controls.Application.Current?.Windows.FirstOrDefault()?.Handler?.PlatformView as Microsoft.UI.Xaml.Window
            ?? throw new InvalidOperationException("A janela da aplicação não está disponível.");
        var picker = new FileSavePicker { SuggestedStartLocation = PickerLocationId.DocumentsLibrary, SuggestedFileName = System.IO.Path.GetFileNameWithoutExtension(suggestedName) };
        picker.FileTypeChoices.Add("Base Excel", new List<string> { ".xlsx" });
        WinRT.Interop.InitializeWithWindow.Initialize(picker, WinRT.Interop.WindowNative.GetWindowHandle(window));
        var file = await picker.PickSaveFileAsync(); if (file is null) return null;
        if (System.IO.Path.GetFullPath(file.Path).Equals(System.IO.Path.GetFullPath(currentPath), StringComparison.OrdinalIgnoreCase))
            throw new IOException("Escolha um nome diferente para a cópia da base.");
        using var transaction = await file.OpenTransactedWriteAsync();
        using (var stream = transaction.Stream.AsStreamForWrite())
        { stream.SetLength(0); await stream.WriteAsync(workbook); await stream.FlushAsync(); await transaction.CommitAsync(); }
        return file.Path;
    }
}
