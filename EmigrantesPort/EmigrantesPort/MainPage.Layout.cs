using EmigrantesPort.Core;
using static EmigrantesPort.Design;

namespace EmigrantesPort;

public partial class MainPage
{
    private View Sidebar()
    {
        var side = Design.Grid("*", "Auto,*,Auto"); side.BackgroundColor = Green; side.Padding = new Thickness(22, 30, 22, 24); side.RowSpacing = 35;
        var brand = new VerticalStackLayout { Spacing = 6 };
        brand.Add(Design.Label("emigrantes", 25, Colors.White, true));
        brand.Add(Design.Label("P O R T U G A L", 10, Color.FromArgb("#ADC6BA")));
        side.Add(brand, 0, 0);
        var navigation = new VerticalStackLayout { Spacing = 9 };
        navigation.Add(Design.Label("ÁREA DE TRABALHO", 10, Color.FromArgb("#ADC6BA"), true));
        foreach (var label in new[] { "Candidatos", "Empresas", "Funções", "Actividades", "Durações", "Revisão" })
        {
            var button = new Button { Text = label, HorizontalOptions = LayoutOptions.Fill, HeightRequest = 44, TextColor = Colors.White, BackgroundColor = Colors.Transparent, Padding = new Thickness(13, 8), FontSize = 13 };
            button.Clicked += (_, _) => { if (busy) return; section = label; page = 0; RefreshAll(); };
            menu.Add(button); navigation.Add(button);
        }
        side.Add(navigation, 0, 1);
        var footer = new VerticalStackLayout { Spacing = 11 };
        footer.Add(fileLabel);
        var reload = new Button { Text = "Reabrir a base", BackgroundColor = Color.FromArgb("#31594C"), TextColor = Colors.White, FontSize = 12 };
        reload.Clicked += async (_, _) => await Run(async () =>
        {
            await Task.Run(session.Reload); RefreshAll(); status.Text = "Base actualizada a partir do Excel.";
        });
        footer.Add(reload);
        var location = new Button { Text = "Localização e segurança", BackgroundColor = Colors.Transparent, TextColor = Color.FromArgb("#ADC6BA"), FontSize = 11, Padding = new Thickness(0) };
        location.Clicked += async (_, _) => await DisplayAlertAsync("A sua base Excel", (session.Path ?? "Base por abrir") + "\n\nAs versões anteriores ficam na pasta CopiasSeguranca, junto da base. Pode restaurar uma versão com «Abrir base».\n\nSe editar as tabelas no Excel, feche o ficheiro e utilize «Reabrir a base». A folha Consulta é actualizada na próxima gravação.", "Fechar");
        footer.Add(location); footer.Add(Design.Label("Uma base. Todas as ligações.", 11, Color.FromArgb("#ADC6BA")));
        side.Add(footer, 0, 2); return side;
    }
    private void BuildWorkspace()
    {
        foreach (var view in new View[] { search, companyFilter, roleFilter, list, status, pageLabel, previous, next })
        {
            if (view.Parent is Layout parent) parent.Remove(view);
            else if (view.Parent is Border border) border.Content = null;
        }
        workspace.Clear();
        var top = Design.Grid("*,Auto,Auto,Auto"); top.ColumnSpacing = 8;
        top.Add(Design.Label("GESTÃO DE TRABALHADORES  /  " + section.ToUpperInvariant(), 11, Muted, true), 0, 0);
        top.Add(Action("Abrir base", OpenDatabase), 1, 0); top.Add(Action("Importar Excel", Import), 2, 0); top.Add(Action("Guardar cópia", Export), 3, 0);
        workspace.Add(top, 0, 0);
        var heading = Design.Grid("*,Auto");
        var titles = new VerticalStackLayout { Spacing = 6 };
        titles.Add(Design.Label(section == "Candidatos" ? "Pessoas, novos caminhos." : section, 30, Ink, true));
        titles.Add(Design.Label(section switch
        {
            "Candidatos" => "Organize os candidatos e acompanhe as suas ligações a empresas em Portugal.",
            "Empresas" => "Empresas contratantes e observações, num só lugar.",
            "Revisão" => "Pontos a confirmar na importação. Os valores originais estão na folha Origem.",
            _ => "Nomes únicos para manter a base consistente e fácil de consultar."
        }, 13, Muted));
        heading.Add(titles, 0, 0);
        if (section != "Revisão") heading.Add(Action(section == "Candidatos" ? "+ Registar candidato" : "+ Adicionar", () => section == "Candidatos" ? EditPlacement(null) : EditCatalog(null), true), 1, 0);
        workspace.Add(heading, 0, 1);
        var stats = Design.Grid("*,*,*"); stats.ColumnSpacing = 14;
        stats.Add(Stat(data.Candidates.Count.ToString(), "Candidatos registados", "Pessoas na base"), 0, 0);
        stats.Add(Stat(data.Companies.Count.ToString(), "Empresas contratantes", "Ligações para Portugal"), 1, 0);
        stats.Add(Stat(data.Placements.Count.ToString(), "Contratações", "Registos de trabalho"), 2, 0);
        workspace.Add(stats, 0, 2);
        var tools = Design.Grid("*,Auto,Auto"); tools.ColumnSpacing = 10;
        tools.Add(Design.Card(search, new Thickness(12, 0)), 0, 0);
        if (section == "Candidatos") { tools.Add(Design.Card(companyFilter, new Thickness(10, 0)), 1, 0); tools.Add(Design.Card(roleFilter, new Thickness(10, 0)), 2, 0); }
        search.Placeholder = section == "Candidatos" ? "Pesquisar nome, empresa, função ou contacto…" : "Pesquisar nesta lista…";
        workspace.Add(tools, 0, 3);
        var table = Design.Grid("*", "Auto,*");
        var headers = Design.Grid(section == "Candidatos" ? "2.3*,2*,1.6*,1*,110" : "2*,3*,110"); headers.Padding = new Thickness(16, 13); headers.BackgroundColor = Color.FromArgb("#EDF2EC"); headers.ColumnSpacing = 12;
        var labels = section == "Candidatos" ? new[] { "CANDIDATO", "EMPRESA", "FUNÇÃO / CONTRATO", "SALÁRIO", "" } : section == "Revisão" ? new[] { "ORIGEM", "A CONFIRMAR", "" } : new[] { "NOME", section == "Empresas" ? "OBSERVAÇÃO" : "DESCRIÇÃO", "" };
        for (var i = 0; i < labels.Length; i++) headers.Add(Design.Label(labels[i], 10, Muted, true), i, 0);
        table.Add(headers, 0, 0);
        list.ItemTemplate = new DataTemplate(() => CreateRow());
        list.EmptyView = new VerticalStackLayout
        {
            Padding = new Thickness(35, 55),
            Spacing = 10,
            Children = { Design.Label("Nada por aqui, por enquanto.", 20, Ink, true), Design.Label("Importe o Excel, adicione um registo ou ajuste os filtros.", 13, Muted) }
        };
        table.Add(list, 0, 1); workspace.Add(Design.Card(table, new Thickness(0)), 0, 4);
        var bottom = Design.Grid("*,Auto,Auto,Auto"); bottom.ColumnSpacing = 10;
        status.LineBreakMode = LineBreakMode.TailTruncation; bottom.Add(status, 0, 0); bottom.Add(pageLabel, 1, 0); bottom.Add(previous, 2, 0); bottom.Add(next, 3, 0);
        workspace.Add(bottom, 0, 5);
        foreach (var button in menu) button.BackgroundColor = button.Text == section ? Color.FromArgb("#31594C") : Colors.Transparent;
    }
    private static View Stat(string value, string label, string detail)
    {
        var layout = new VerticalStackLayout { Spacing = 4 }; layout.Add(Design.Label(label, 12, Muted));
        layout.Add(Design.Label(value, 30, Ink, true)); layout.Add(Design.Label(detail, 11, Muted)); return Design.Card(layout, new Thickness(20, 15));
    }
    private View CreateRow()
    {
        var grid = Design.Grid(section == "Candidatos" ? "2.3*,2*,1.6*,1*,110" : "2*,3*,110"); grid.Padding = new Thickness(16, 12); grid.ColumnSpacing = 12;
        Label Bound(string property, double size = 13, bool bold = false, Color? color = null)
        { var l = Design.Label("", size, color, bold); l.LineBreakMode = LineBreakMode.TailTruncation; l.SetBinding(Microsoft.Maui.Controls.Label.TextProperty, property); return l; }
        if (section == "Candidatos")
        {
            grid.Add(new VerticalStackLayout { Spacing = 5, Children = { Bound(nameof(DisplayRow.Name), 13, true), Bound(nameof(DisplayRow.Contact), 11, false, Muted) } }, 0, 0);
            grid.Add(Bound(nameof(DisplayRow.Company)), 1, 0);
            grid.Add(new VerticalStackLayout { Spacing = 5, Children = { Bound(nameof(DisplayRow.Role)), Bound(nameof(DisplayRow.Duration), 11, false, Muted) } }, 2, 0);
            grid.Add(Bound(nameof(DisplayRow.Salary), 13, true), 3, 0);
        }
        else { grid.Add(Bound(nameof(DisplayRow.Name), 13, true), 0, 0); grid.Add(Bound(nameof(DisplayRow.Description)), 1, 0); }
        if (section != "Revisão")
        {
            var actions = new HorizontalStackLayout { Spacing = 3, HorizontalOptions = LayoutOptions.End };
            var edit = new Button { Text = "Editar", BackgroundColor = Colors.Transparent, TextColor = Green, Padding = new Thickness(6), FontSize = 11 };
            edit.Clicked += async (sender, _) => await Run(async () => { var row = (DisplayRow)((Button)sender!).BindingContext; if (section == "Candidatos") await EditPlacement(row.Id); else await EditCatalog(row.Id); });
            var remove = new Button { Text = "×", BackgroundColor = Colors.Transparent, TextColor = Color.FromArgb("#9C5447"), Padding = new Thickness(9), FontSize = 20 };
            SemanticProperties.SetDescription(remove, "Remover registo");
            remove.Clicked += async (sender, _) => await Run(() => Remove((DisplayRow)((Button)sender!).BindingContext));
            actions.Add(edit); actions.Add(remove); grid.Add(actions, section == "Candidatos" ? 4 : 2, 0);
        }
        return new Border { Content = grid, Stroke = Line, StrokeThickness = 0.5, BackgroundColor = Colors.White };
    }
}
