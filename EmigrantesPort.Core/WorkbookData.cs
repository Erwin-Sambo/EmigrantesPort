using System.Globalization;

namespace EmigrantesPort.Core;

public sealed class WorkbookData
{
    public List<Candidate> Candidates { get; } = [];
    public List<Company> Companies { get; } = [];
    public List<Lookup> Roles { get; } = [];
    public List<Lookup> Activities { get; } = [];
    public List<Lookup> Durations { get; } = [];
    public List<Placement> Placements { get; } = [];
    public List<SourceRow> Sources { get; } = [];
    public List<ImportIssue> Issues { get; } = [];
    public WorkbookData Copy()
    {
        var copy = new WorkbookData();
        copy.Candidates.AddRange(Candidates); copy.Companies.AddRange(Companies);
        copy.Roles.AddRange(Roles); copy.Activities.AddRange(Activities); copy.Durations.AddRange(Durations);
        copy.Placements.AddRange(Placements); copy.Sources.AddRange(Sources); copy.Issues.AddRange(Issues);
        return copy;
    }
    internal void ReplaceWith(WorkbookData other)
    {
        static void Replace<T>(List<T> target, List<T> source)
        {
            target.Clear();
            target.AddRange(source);
        }

        Replace(Candidates, other.Candidates);
        Replace(Companies, other.Companies);
        Replace(Roles, other.Roles);
        Replace(Activities, other.Activities);
        Replace(Durations, other.Durations);
        Replace(Placements, other.Placements);
        Replace(Sources, other.Sources);
        Replace(Issues, other.Issues);
    }

    public IReadOnlyList<PlacementView> Query(string search = "", string? companyId = null, string? roleId = null)
    {
        var candidates = Candidates.ToDictionary(x => x.Id); var companies = Companies.ToDictionary(x => x.Id);
        var roles = Roles.ToDictionary(x => x.Id); var durations = Durations.ToDictionary(x => x.Id);
        var activities = Activities.ToDictionary(x => x.Id);
        var key = Text.Key(search);
        return Placements.Where(p => (companyId is null || p.CompanyId == companyId) && (roleId is null || p.RoleId == roleId))
            .Select(p => new PlacementView(p.Id, p.CandidateId, candidates[p.CandidateId].Name,
                companies.GetValueOrDefault(p.CompanyId)?.Name ?? "Por indicar", roles.GetValueOrDefault(p.RoleId)?.Name ?? "Por indicar",
                durations.GetValueOrDefault(p.DurationId)?.Name ?? "Por indicar", activities.GetValueOrDefault(p.ActivityId)?.Name ?? "Por indicar",
                p.Salary, candidates[p.CandidateId].Contact, p.Observation))
            .Where(p => key.Length == 0 || Text.Key(string.Join(" ", p.Name, p.Contact, p.Company, p.Role, p.Activity, p.Duration, p.Observation)).Contains(key))
            .OrderBy(p => p.Name, StringComparer.Create(CultureInfo.GetCultureInfo("pt-PT"), true)).ToList();
    }
    public List<Lookup> LookupList(Catalog catalog) => catalog switch
    { Catalog.Roles => Roles, Catalog.Activities => Activities, Catalog.Durations => Durations, _ => throw new ArgumentException("Seleccione um catálogo.") };
    public string AddCatalog(Catalog catalog, string name, string description)
    {
        name = Text.Clean(name); description = description.Trim();
        if (name.Length == 0) throw new InvalidDataException("O nome é obrigatório.");
        if (catalog == Catalog.Companies)
        {
            if (Companies.Any(x => Text.Key(x.Name) == Text.Key(name))) throw new InvalidDataException("Já existe uma empresa com esse nome.");
            var id = Text.Id(); Companies.Add(new(id, name, description)); return id;
        }
        var list = LookupList(catalog);
        if (list.Any(x => Text.Key(x.Name) == Text.Key(name))) throw new InvalidDataException("Já existe um registo com esse nome.");
        var key = Text.Id(); list.Add(new(key, name, description)); return key;
    }
    public void EditCatalog(Catalog catalog, string id, string name, string description)
    {
        name = Text.Clean(name);
        if (name.Length == 0) throw new InvalidDataException("O nome é obrigatório.");
        if (catalog == Catalog.Companies)
        {
            if (Companies.Any(x => x.Id != id && Text.Key(x.Name) == Text.Key(name))) throw new InvalidDataException("O nome da empresa deve ser único.");
            var index = Companies.FindIndex(x => x.Id == id); if (index < 0) throw new InvalidDataException("Empresa inexistente.");
            Companies[index] = new(id, name, description.Trim()); return;
        }
        var list = LookupList(catalog);
        if (list.Any(x => x.Id != id && Text.Key(x.Name) == Text.Key(name))) throw new InvalidDataException("O nome deve ser único.");
        var i = list.FindIndex(x => x.Id == id); if (i < 0) throw new InvalidDataException("Registo inexistente.");
        list[i] = new(id, name, description.Trim());
    }
    public void DeleteCatalog(Catalog catalog, string id)
    {
        if (Placements.Any(p => catalog switch
        {
            Catalog.Companies => p.CompanyId == id,
            Catalog.Roles => p.RoleId == id,
            Catalog.Activities => p.ActivityId == id,
            _ => p.DurationId == id
        }))
            throw new InvalidDataException("Este registo está ligado a contratações. Altere essas ligações antes de o remover.");
        if (catalog == Catalog.Companies) Companies.RemoveAll(x => x.Id == id); else LookupList(catalog).RemoveAll(x => x.Id == id);
    }
    public string SavePlacement(PlacementInput input, string? id = null)
    {
        var name = Text.Clean(input.Name);
        if (name.Length == 0) throw new InvalidDataException("Indique o nome do candidato.");
        if (!Companies.Any(x => x.Id == input.CompanyId)) throw new InvalidDataException("Seleccione uma empresa.");
        if (!Roles.Any(x => x.Id == input.RoleId)) throw new InvalidDataException("Seleccione uma função.");
        if (!Durations.Any(x => x.Id == input.DurationId)) throw new InvalidDataException("Seleccione a duração do contrato.");
        if (!Activities.Any(x => x.Id == input.ActivityId)) throw new InvalidDataException("Seleccione uma actividade.");
        if (input.Salary < 0) throw new InvalidDataException("O salário não pode ser negativo.");
        var existing = id is null ? null : Placements.SingleOrDefault(x => x.Id == id) ?? throw new InvalidDataException("Contratação inexistente.");
        string candidateId;
        if (existing is null)
        {
            if (!string.IsNullOrEmpty(input.ExistingCandidateId))
            {
                if (!Candidates.Any(c => c.Id == input.ExistingCandidateId)) throw new InvalidDataException("Candidato inexistente.");
                candidateId = input.ExistingCandidateId;
            }
            else { candidateId = Text.Id(); Candidates.Add(new(candidateId, name, Text.Clean(input.Contact))); }
        }
        else
        {
            candidateId = existing.CandidateId;
            var c = Candidates.FindIndex(x => x.Id == candidateId);
            Candidates[c] = new(candidateId, name, Text.Clean(input.Contact));
        }
        var value = new Placement(id ?? Text.Id(), candidateId, input.CompanyId, input.RoleId, input.DurationId, input.ActivityId, input.Salary, input.Observation.Trim());
        if (existing is null) Placements.Add(value); else Placements[Placements.IndexOf(existing)] = value;
        return value.Id;
    }
    public void DeletePlacement(string id)
    {
        var p = Placements.SingleOrDefault(x => x.Id == id) ?? throw new InvalidDataException("Contratação inexistente.");
        Placements.Remove(p); Sources.RemoveAll(x => x.PlacementId == id);
        if (!Placements.Any(x => x.CandidateId == p.CandidateId)) Candidates.RemoveAll(x => x.Id == p.CandidateId);
    }
}
