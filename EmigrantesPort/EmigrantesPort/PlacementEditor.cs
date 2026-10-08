using EmigrantesPort.Core;
using static EmigrantesPort.Design;

namespace EmigrantesPort;

internal sealed class PlacementEditor : ContentPage
{
    private readonly Entry name = new() { Placeholder = "Nome completo" }, contact = new() { Placeholder = "Contacto telefónico", Keyboard = Keyboard.Telephone }, salary = new() { Placeholder = "Ex.: 920,00", Keyboard = Keyboard.Numeric };
    private readonly Picker companies = new() { Title = "Seleccionar empresa" }, roles = new() { Title = "Seleccionar função" }, durations = new() { Title = "Seleccionar duração" }, activities = new() { Title = "Seleccionar actividade" };
    private readonly Editor observation = new() { Placeholder = "Observação deste registo (opcional)", AutoSize = EditorAutoSizeOption.TextChanges, MinimumHeightRequest = 75 };
    public PlacementEditor(WorkbookData data, string? id, Func<PlacementInput, Task> save)
    {
        Title = id is null ? "Registar candidato" : "Editar contratação"; BackgroundColor = Canvas;
        var existingCandidate = new Picker { Title = "Novo candidato" };
        existingCandidate.ItemsSource = new List<Choice> { new("", "Registar uma nova pessoa") }.Concat(data.Candidates.OrderBy(c => c.Name).Select(c => new Choice(c.Id, c.Name + (c.Contact.Length == 0 ? " · sem contacto" : " · " + c.Contact)))).ToList();
        existingCandidate.SelectedIndex = 0;
        existingCandidate.SelectedIndexChanged += (_, _) =>
        {
            var selectedId = (existingCandidate.SelectedItem as Choice)?.Id;
            var selected = data.Candidates.FirstOrDefault(c => c.Id == selectedId);
            name.Text = selected?.Name ?? ""; contact.Text = selected?.Contact ?? "";
            name.IsReadOnly = selected is not null; contact.IsReadOnly = selected is not null;
        };
        companies.ItemsSource = data.Companies.OrderBy(c => c.Name).Select(c => new Choice(c.Id, c.Name)).ToList();
        roles.ItemsSource = data.Roles.OrderBy(c => c.Name).Select(c => new Choice(c.Id, c.Name)).ToList();
        durations.ItemsSource = data.Durations.OrderBy(c => c.Name).Select(c => new Choice(c.Id, c.Name)).ToList();
        activities.ItemsSource = data.Activities.OrderBy(c => c.Name).Select(c => new Choice(c.Id, c.Name)).ToList();
        if (id is not null)
        {
            var p = data.Placements.Single(p => p.Id == id); var c = data.Candidates.Single(c => c.Id == p.CandidateId);
            name.Text = c.Name; contact.Text = c.Contact; salary.Text = p.Salary?.ToString("0.00", System.Globalization.CultureInfo.GetCultureInfo("pt-PT")); observation.Text = p.Observation;
            Select(companies, p.CompanyId); Select(roles, p.RoleId); Select(durations, p.DurationId); Select(activities, p.ActivityId);
        }
        var layout = new VerticalStackLayout { Spacing = 22, Padding = new Thickness(35, 30), MaximumWidthRequest = 950, HorizontalOptions = LayoutOptions.Fill };
        layout.Add(Design.Label(Title, 30, Ink, true)); layout.Add(Design.Label("Dados claros, ligações certas. Os campos com * são obrigatórios.", 13, Muted));
        var person = new VerticalStackLayout { Spacing = 18 };
        person.Add(Design.Label("01  ·  O candidato", 16, Ink, true));
        if (id is null) person.Add(Field("Pessoa a ligar à empresa", existingCandidate, "Seleccione uma pessoa existente para registar outro contrato sem duplicar o candidato."));
        var names = Design.Grid("2*,*"); names.ColumnSpacing = 18; names.Add(Field("Nome do candidato *", name), 0, 0); names.Add(Field("Contacto", contact), 1, 0); person.Add(names);
        if (id is not null) person.Add(Design.Label("O nome e o contacto são partilhados por todas as contratações deste candidato.", 11, Muted));
        layout.Add(Card(person));
        var job = new VerticalStackLayout { Spacing = 18 }; job.Add(Design.Label("02  ·  Trabalho em Portugal", 16, Ink, true));
        var line1 = Design.Grid("*,*"); line1.ColumnSpacing = 18; line1.Add(Field("Empresa contratante *", companies), 0, 0); line1.Add(Field("Função a exercer *", roles), 1, 0); job.Add(line1);
        var line2 = Design.Grid("*,*"); line2.ColumnSpacing = 18; line2.Add(Field("Duração do contrato *", durations), 0, 0); line2.Add(Field("Actividade *", activities), 1, 0); job.Add(line2);
        job.Add(Field("Salário mensal (€)", salary, "Pode deixar vazio quando o salário ainda não estiver confirmado."));
        job.Add(Field("Observação original / do registo", observation)); layout.Add(Card(job));
        var error = Design.Label("", 13, Color.FromArgb("#9C5447")); error.IsVisible = false; layout.Add(error);
        var buttons = new HorizontalStackLayout { Spacing = 12, HorizontalOptions = LayoutOptions.End };
        var cancel = Design.Button("Cancelar"); cancel.Clicked += async (_, _) => await Navigation.PopAsync();
        var submit = Design.Button("Guardar no Excel", true);
        submit.Clicked += async (_, _) =>
        {
            submit.IsEnabled = false; cancel.IsEnabled = false;
            try
            {
                string Selected(Picker picker) => (picker.SelectedItem as Choice)?.Id ?? "";
                await save(new(name.Text ?? "", contact.Text ?? "", Selected(companies), Selected(roles), Selected(durations), Selected(activities), Text.Salary(salary.Text ?? ""), observation.Text ?? "", id is null ? (existingCandidate.SelectedItem as Choice)?.Id : null));
                await Navigation.PopAsync();
            }
            catch (Exception ex) { error.Text = MainPage.Friendly(ex); error.IsVisible = true; }
            finally { submit.IsEnabled = true; cancel.IsEnabled = true; }
        };
        buttons.Add(cancel); buttons.Add(submit); layout.Add(buttons); Content = new ScrollView { Content = layout };
    }
    private static void Select(Picker picker, string id) => picker.SelectedItem = ((IEnumerable<Choice>)picker.ItemsSource).FirstOrDefault(c => c.Id == id);
}
