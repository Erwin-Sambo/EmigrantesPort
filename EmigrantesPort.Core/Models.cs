using System.Globalization;

namespace EmigrantesPort.Core;

public sealed record Candidate(string Id, string Name, string Contact);
public sealed record Company(string Id, string Name, string Observation);
public sealed record Lookup(string Id, string Name, string Description);
public sealed record Placement(string Id, string CandidateId, string CompanyId, string RoleId,
    string DurationId, string ActivityId, decimal? Salary, string Observation);
public sealed record SourceRow(string Id, string PlacementId, string File, string Sheet, int Row, string OriginalJson);
public sealed record ImportIssue(string Id, string Source, string Message);
public sealed record PlacementView(string Id, string CandidateId, string Name, string Company, string Role,
    string Duration, string Activity, decimal? Salary, string Contact, string Observation)
{
    public string SalaryText => Salary is null ? "Por indicar" : Salary.Value.ToString("N2", CultureInfo.GetCultureInfo("pt-PT")) + " €";
}
public sealed record PlacementInput(string Name, string Contact, string CompanyId, string RoleId,
    string DurationId, string ActivityId, decimal? Salary, string Observation = "", string? ExistingCandidateId = null);
public sealed record ImportResult(int Imported, int Duplicates, int Issues, int Unassigned = 0);
public enum Catalog { Companies, Roles, Activities, Durations }
