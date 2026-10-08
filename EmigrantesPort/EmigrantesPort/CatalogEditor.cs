using EmigrantesPort.Core;
using static EmigrantesPort.Design;

namespace EmigrantesPort;

internal sealed class CatalogEditor : ContentPage
{
    public CatalogEditor(string section, bool company, string value, string detail, Func<string, string, Task> save)
    {
        Title = value.Length == 0 ? "Adicionar · " + section : "Editar · " + section; BackgroundColor = Canvas;
        var name = new Entry { Text = value, Placeholder = company ? "Nome da empresa" : "Nome" };
        var description = new Editor { Text = detail, Placeholder = company ? "Ex.: Fund. 2025" : "Descrição (opcional)", AutoSize = EditorAutoSizeOption.TextChanges, MinimumHeightRequest = 100 };
        var layout = new VerticalStackLayout { Spacing = 22, Padding = new Thickness(35, 30), MaximumWidthRequest = 800 };
        layout.Add(Design.Label(Title, 28, Ink, true)); layout.Add(Design.Label("O nome é único. Alterações aparecem em todas as contratações ligadas.", 13, Muted));
        layout.Add(Card(new VerticalStackLayout { Spacing = 20, Children = { Field("Nome *", name), Field(company ? "Observação da empresa" : "Descrição", description) } }));
        var error = Design.Label("", 13, Color.FromArgb("#9C5447")); error.IsVisible = false; layout.Add(error);
        var buttons = new HorizontalStackLayout { Spacing = 12, HorizontalOptions = LayoutOptions.End };
        var cancel = Design.Button("Cancelar"); cancel.Clicked += async (_, _) => await Navigation.PopAsync();
        var submit = Design.Button("Guardar no Excel", true); submit.Clicked += async (_, _) =>
        {
            submit.IsEnabled = false; cancel.IsEnabled = false;
            try { await save(name.Text ?? "", description.Text ?? ""); await Navigation.PopAsync(); }
            catch (Exception ex) { error.Text = MainPage.Friendly(ex); error.IsVisible = true; }
            finally { submit.IsEnabled = true; cancel.IsEnabled = true; }
        };
        buttons.Add(cancel); buttons.Add(submit); layout.Add(buttons); Content = new ScrollView { Content = layout };
    }
}
