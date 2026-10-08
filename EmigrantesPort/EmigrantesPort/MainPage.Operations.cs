using EmigrantesPort.Core;
using static EmigrantesPort.Design;

namespace EmigrantesPort;

public partial class MainPage
{
    private async Task<FileResult?> Pick(string title) => await FilePicker.Default.PickAsync(new PickOptions
    {
        PickerTitle = title,
        FileTypes = new FilePickerFileType(new Dictionary<DevicePlatform, IEnumerable<string>> { [DevicePlatform.WinUI] = [".xlsx"], [DevicePlatform.macOS] = ["xlsx"], [DevicePlatform.Android] = ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"] })
    });
    private async Task OpenDatabase()
    {
        var file = await Pick("Abrir base Excel normalizada"); if (file is null) return;
        await Task.Run(() => session.Open(file.FullPath));
        initialized = true;
        Preferences.Default.Set("workbook.path", session.Path!);
        page = 0;
        RefreshAll(); status.Text = "Base aberta: " + session.Path;
    }
    private async Task Import()
    {
        if (!initialized || !session.IsOpen) throw new InvalidOperationException("Abra a base antes de importar.");
        var file = await Pick("Importar candidatos do Excel original"); if (file is null) return;
        status.Text = "A analisar e normalizar o Excel…";
        var prepared = await Task.Run(() => session.PrepareImport(file.FullPath));
        var result = prepared.Result;
        var proceed = await DisplayAlertAsync("Importação preparada", $"{result.Imported} novas contratações\n{result.Duplicates} linhas repetidas\n{result.Unassigned} linhas sem candidato (preservadas)\n{result.Issues} pontos para revisão\n\nAs linhas originais serão preservadas em Origem. O ficheiro seleccionado não será alterado. Importar para a base actual?", "Importar", "Cancelar");
        if (!proceed) { status.Text = "Importação cancelada. A base não foi alterada."; return; }
        await Task.Run(() => session.CommitImport(prepared)); page = 0; RefreshAll();
        status.Text = $"Importadas {result.Imported} contratações. Consulte a área Revisão.";
    }
    private async Task Export()
    {
        if (!initialized || !session.IsOpen) throw new InvalidOperationException("Abra a base antes de exportar.");
        var bytes = await Task.Run(() =>
        {
            using var book = ExcelStore.Build(data); using var stream = new MemoryStream(); book.SaveAs(stream); return stream.ToArray();
        });
        var path = await destination.SaveCopyAsync(bytes, $"EmigrantesPort-{DateTime.Now:yyyy-MM-dd}.xlsx", session.Path!);
        if (path is not null) status.Text = "Cópia guardada: " + path;
    }
    private async Task Commit(Action<WorkbookData> mutation)
    {
        await Task.Run(() => session.Apply(mutation));
        RefreshAll();
        status.Text = "Alterações guardadas no Excel. Cópia de segurança criada.";
    }
    private async Task Remove(DisplayRow row)
    {
        if (!await DisplayAlertAsync("Remover registo", $"{(section == "Candidatos" ? "Remover esta contratação de" : "Remover")} «{row.Name}»?\nA versão anterior da base fica na pasta CopiasSeguranca.", "Remover", "Cancelar")) return;
        await Commit(copy => { if (section == "Candidatos") copy.DeletePlacement(row.Id); else copy.DeleteCatalog(CurrentCatalog(), row.Id); });
    }
    private async Task EditPlacement(string? id)
    {
        if (data.Companies.Count == 0 || data.Roles.Count == 0 || data.Activities.Count == 0 || data.Durations.Count == 0)
        { await DisplayAlertAsync("Prepare os catálogos", "Adicione pelo menos uma empresa, função, actividade e duração nas áreas à esquerda, ou importe o Excel existente.", "Entendido"); return; }
        var editor = new PlacementEditor(data, id, input => Commit(copy => copy.SavePlacement(input, id)));
        await Navigation.PushAsync(editor);
    }
    private async Task EditCatalog(string? id)
    {
        var catalog = CurrentCatalog();
        string name = "", description = "";
        if (id is not null)
        {
            if (catalog == Catalog.Companies) { var item = data.Companies.Single(c => c.Id == id); name = item.Name; description = item.Observation; }
            else { var item = data.LookupList(catalog).Single(c => c.Id == id); name = item.Name; description = item.Description; }
        }
        var editor = new CatalogEditor(section, catalog == Catalog.Companies, name, description, async (n, d) =>
        { await Commit(copy => { if (id is null) copy.AddCatalog(catalog, n, d); else copy.EditCatalog(catalog, id, n, d); }); });
        await Navigation.PushAsync(editor);
    }
}
