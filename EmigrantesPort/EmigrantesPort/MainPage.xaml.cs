using EmigrantesPort.Core;
using static EmigrantesPort.Design;

namespace EmigrantesPort;

public partial class MainPage : ContentPage
{
    private readonly IWorkbookDestination destination;
    private readonly WorkbookSession session = new();
    private WorkbookData data => session.Data;
    private bool initialized, busy, refreshing;
    private string section = "Candidatos";
    private readonly Grid workspace = Design.Grid("*", "Auto,Auto,Auto,Auto,*,Auto");
    private readonly Label status = Design.Label("A preparar a base Excel…", 11, Muted);
    private readonly Label fileLabel = Design.Label("BASE EXCEL", 11, Color.FromArgb("#ADC6BA"));
    private readonly Entry search = new() { Placeholder = "Pesquisar nome, empresa, função ou contacto…", ClearButtonVisibility = ClearButtonVisibility.WhileEditing };
    private readonly Picker companyFilter = new() { Title = "Todas as empresas", WidthRequest = 210 };
    private readonly Picker roleFilter = new() { Title = "Todas as funções", WidthRequest = 180 };
    private readonly Label pageLabel = Design.Label("", 12, Muted);
    private readonly CollectionView list = new() { SelectionMode = SelectionMode.None, ItemsLayout = new LinearItemsLayout(ItemsLayoutOrientation.Vertical) { ItemSpacing = 0 } };
    private readonly List<Button> menu = [];
    private readonly Button previous = Design.Button("Anterior"), next = Design.Button("Seguinte");
    private List<Choice> companyChoices = [], roleChoices = [];
    private int page;
    private const int PageSize = 25;
    public MainPage(IWorkbookDestination destination)
    {
        InitializeComponent(); this.destination = destination;
        NavigationPage.SetHasNavigationBar(this, false);
        var root = Design.Grid("218,*", "*"); root.BackgroundColor = Canvas;
        var sidebar = Sidebar(); root.Add(sidebar, 0, 0);
        workspace.Padding = new Thickness(30, 24, 30, 18); workspace.RowSpacing = 18;
        root.Add(workspace, 1, 0); Content = root;
        search.TextChanged += (_, _) => { if (!refreshing) { page = 0; RefreshList(); } };
        companyFilter.SelectedIndexChanged += (_, _) => { if (!refreshing) { page = 0; RefreshList(); } };
        roleFilter.SelectedIndexChanged += (_, _) => { if (!refreshing) { page = 0; RefreshList(); } };
        previous.Clicked += (_, _) => { page--; RefreshList(); };
        next.Clicked += (_, _) => { page++; RefreshList(); };
        BuildWorkspace();
    }
    protected override async void OnAppearing()
    {
        base.OnAppearing(); if (initialized) return;
        await Run(async () =>
        {
            var path = Preferences.Default.Get("workbook.path", System.IO.Path.Combine(FileSystem.AppDataDirectory, "EmigrantesPort", "Base.xlsx"));
            var createIfMissing = !Preferences.Default.ContainsKey("workbook.path");
            await Task.Run(() => session.Open(path, createIfMissing));
            initialized = true; RefreshAll(); status.Text = "Base aberta. As alterações são guardadas automaticamente no Excel.";
        });
    }
    private Button Action(string text, Func<Task> action, bool primary = false)
    { var button = Design.Button(text, primary); button.Clicked += async (_, _) => await Run(action); return button; }
    private async Task Run(Func<Task> action)
    {
        if (busy) return; busy = true;
        try { workspace.IsEnabled = false; await action(); }
        catch (Exception ex) { status.Text = "Não foi possível concluir a operação."; await DisplayAlertAsync("Operação não concluída", Friendly(ex), "Fechar"); }
        finally { busy = false; workspace.IsEnabled = true; }
    }
    internal static string Friendly(Exception ex) => ex is IOException ? ex.Message + "\nSe o ficheiro estiver aberto no Excel, feche-o e tente novamente." : ex.Message;
    private void RefreshAll()
    {
        refreshing = true;
        var selectedCompany = (companyFilter.SelectedItem as Choice)?.Id; var selectedRole = (roleFilter.SelectedItem as Choice)?.Id;
        companyChoices = [new("", "Todas as empresas"), .. data.Companies.OrderBy(c => c.Name).Select(c => new Choice(c.Id, c.Name))];
        roleChoices = [new("", "Todas as funções"), .. data.Roles.OrderBy(c => c.Name).Select(c => new Choice(c.Id, c.Name))];
        companyFilter.ItemsSource = companyChoices; companyFilter.SelectedItem = companyChoices.FirstOrDefault(c => c.Id == selectedCompany) ?? companyChoices[0];
        roleFilter.ItemsSource = roleChoices; roleFilter.SelectedItem = roleChoices.FirstOrDefault(c => c.Id == selectedRole) ?? roleChoices[0];
        refreshing = false;
        fileLabel.Text = "BASE EXCEL\n" + (session.Path is null ? "Por abrir" : System.IO.Path.GetFileName(session.Path));
        SemanticProperties.SetDescription(fileLabel, session.Path ?? "Base por abrir");
        BuildWorkspace(); RefreshList();
    }
    private void RefreshList()
    {
        if (!initialized) return;
        List<DisplayRow> rows;
        if (section == "Candidatos") rows = data.Query(search.Text ?? "", Empty((companyFilter.SelectedItem as Choice)?.Id), Empty((roleFilter.SelectedItem as Choice)?.Id))
            .Select(p => new DisplayRow(p.Id, p.Name, p.Contact.Length == 0 ? "Contacto por indicar" : p.Contact, p.Company, p.Role, p.Duration, p.SalaryText, "")).ToList();
        else
        {
            rows = section switch
            {
                "Empresas" => data.Companies.Select(c => new DisplayRow(c.Id, c.Name, "", "", "", "", "", c.Observation)).ToList(),
                "Revisão" => data.Issues.Select(c => new DisplayRow(c.Id, c.Source, "", "", "", "", "", c.Message)).ToList(),
                _ => data.LookupList(CurrentCatalog()).Select(c => new DisplayRow(c.Id, c.Name, "", "", "", "", "", c.Description)).ToList()
            };
            var key = Text.Key(search.Text); rows = rows.Where(r => Text.Key(r.Name + " " + r.Description).Contains(key)).OrderBy(r => r.Name).ToList();
        }
        var pages = Math.Max(1, (rows.Count + PageSize - 1) / PageSize); page = Math.Clamp(page, 0, pages - 1);
        list.ItemsSource = rows.Skip(page * PageSize).Take(PageSize).ToList();
        pageLabel.Text = $"{rows.Count} registos · {page + 1}/{pages}"; previous.IsEnabled = page > 0; next.IsEnabled = page + 1 < pages;
    }
    private static string? Empty(string? value) => string.IsNullOrEmpty(value) ? null : value;
    private Catalog CurrentCatalog() => section switch { "Empresas" => Catalog.Companies, "Funções" => Catalog.Roles, "Actividades" => Catalog.Activities, _ => Catalog.Durations };
    private sealed record DisplayRow(string Id, string Name, string Contact, string Company, string Role, string Duration, string Salary, string Description);
}
