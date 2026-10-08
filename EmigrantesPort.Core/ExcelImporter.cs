using ClosedXML.Excel;
using System.Globalization;
using System.Text.Json;

namespace EmigrantesPort.Core;

public static class ExcelImporter
{
    public static ImportResult Import(WorkbookData data, string path)
    {
        ExcelStore.Validate(data);
        var workingCopy = data.Copy();
        var result = ImportInto(workingCopy, path);
        data.ReplaceWith(workingCopy);
        return result;
    }

    private static ImportResult ImportInto(WorkbookData data, string path)
    {
        using var book = new XLWorkbook(path);
        if (book.TryGetWorksheet("Configuracao", out var config) && config.Cell(1, 1).GetString() == "EmigrantesPort")
            throw new InvalidDataException("Esta é uma base normalizada. Utilize «Abrir base».");
        var imported = 0; var duplicates = 0; var unassigned = 0; var issuesBefore = data.Issues.Count; var found = false;
        var candidates = data.Candidates.ToDictionary(c => c.Id);
        var companies = data.Companies.ToDictionary(c => c.Id);
        var roles = data.Roles.ToDictionary(c => c.Id); var activities = data.Activities.ToDictionary(c => c.Id); var durations = data.Durations.ToDictionary(c => c.Id);
        string Signature(string name, string contact, string company, string role, string duration, string activity, decimal? salary, string observation, string invalidSalary = "") =>
            JsonSerializer.Serialize(new[] { Text.Key(name), Text.Key(contact), Text.Key(company), Text.Key(role), Text.Key(duration), Text.Key(activity), salary?.ToString(CultureInfo.InvariantCulture) ?? "", Text.Key(observation), Text.Key(invalidSalary) });
        var signatures = data.Placements.GroupBy(p => Signature(candidates[p.CandidateId].Name, candidates[p.CandidateId].Contact,
            companies.GetValueOrDefault(p.CompanyId)?.Name ?? "", roles.GetValueOrDefault(p.RoleId)?.Name ?? "", durations.GetValueOrDefault(p.DurationId)?.Name ?? "",
            activities.GetValueOrDefault(p.ActivityId)?.Name ?? "", p.Salary, p.Observation)).ToDictionary(g => g.Key, g => g.First().Id, StringComparer.Ordinal);
        var contactCandidates = data.Candidates.Where(c => c.Contact.Length > 0).GroupBy(c => Text.Key(c.Name) + "\n" + Text.Key(c.Contact))
            .Where(g => g.Count() == 1).ToDictionary(g => g.Key, g => g.First().Id);
        var companyNames = data.Companies.ToDictionary(c => Text.Key(c.Name), c => c.Id);
        var roleNames = data.Roles.ToDictionary(c => Text.Key(c.Name), c => c.Id);
        var durationNames = data.Durations.ToDictionary(c => Text.Key(c.Name), c => c.Id);
        var activityNames = data.Activities.ToDictionary(c => Text.Key(c.Name), c => c.Id);
        var provenance = data.Sources.Select(s => JsonSerializer.Serialize(new object[] { s.File, s.Sheet, s.Row, s.OriginalJson })).ToHashSet();
        foreach (var sheet in book.Worksheets)
        {
            var lastRow = sheet.LastRowUsed(XLCellsUsedOptions.Contents)?.RowNumber() ?? 0;
            var lastColumn = sheet.LastColumnUsed(XLCellsUsedOptions.Contents)?.ColumnNumber() ?? 0;
            var header = 0; Dictionary<string, int>? columns = null;
            for (var r = 1; r <= Math.Min(lastRow, 20); r++)
            {
                columns = sheet.Row(r).Cells(1, Math.Min(lastColumn, 200)).GroupBy(c => Text.Key(Value(c))).ToDictionary(g => g.Key, g => g.First().Address.ColumnNumber);
                if (columns.ContainsKey("NOME DO CANDIDATO")) { header = r; break; }
            }
            if (header == 0) continue;
            if (lastRow > 100001 || lastColumn > 200) throw new InvalidDataException($"A folha {sheet.Name} excede o limite de importação.");
            found = true;
            int Column(string name, int fallback) => columns!.GetValueOrDefault(Text.Key(name), fallback);
            for (var r = header + 1; r <= lastRow; r++)
            {
                string Read(string name, int fallback) => Text.Clean(Value(sheet.Cell(r, Column(name, fallback))));
                var name = Read("Nome do Candidato", 2);
                if (name.Length == 0 && !Enumerable.Range(2, Math.Min(lastColumn, 9) - 1).Any(c => Text.Clean(Value(sheet.Cell(r, c))).Length > 0)) continue;
                var company = Read("Empresa Contratante", 3); var role = Read("Função a exercer", 4);
                var duration = Text.Duration(Read("Duração do Contrato", Column("Duração do", 5)));
                var activity = Read("Actividade", 6); var contact = Read("Contacto", 8); var observation = Read("Observação", 9);
                var raw = JsonSerializer.Serialize(Enumerable.Range(1, lastColumn).Select(c => Value(sheet.Cell(r, c))).ToArray());
                var source = $"{System.IO.Path.GetFileName(path)} / {sheet.Name} / linha {r}";
                var sourceKey = JsonSerializer.Serialize(new object[] { System.IO.Path.GetFileName(path), sheet.Name, r, raw });
                if (provenance.Contains(sourceKey)) { duplicates++; continue; }
                void Issue(string message) => data.Issues.Add(new(Text.Id(), source, message));
                if (name.Length == 0)
                {
                    data.Sources.Add(new(Text.Id(), "", System.IO.Path.GetFileName(path), sheet.Name, r, raw)); provenance.Add(sourceKey);
                    Issue("Linha sem nome de candidato. Valores preservados em Origem; confirme se é um registo ou uma linha de modelo.");
                    unassigned++; continue;
                }
                decimal? salary = null;
                var invalidSalary = "";
                var salaryCell = sheet.Cell(r, Column("Salário", 7));
                try
                {
                    salary = salaryCell.DataType == XLDataType.Number && !salaryCell.HasFormula ? salaryCell.GetValue<decimal>() : Text.Salary(Value(salaryCell));
                    if (salary < 0) throw new InvalidDataException("Salário negativo.");
                }
                catch (Exception ex) when (ex is InvalidDataException or FormatException or OverflowException)
                { salary = null; invalidSalary = Value(salaryCell); Issue("Salário por confirmar; valor original preservado: " + Value(salaryCell)); }
                if (contact.Length == 0) Issue("Contacto por indicar.");
                if (company.Length == 0 || role.Length == 0 || duration.Length == 0 || activity.Length == 0) Issue("Existem dados de contratação por indicar.");
                if (salary is null) Issue("Salário por indicar.");
                if (lastColumn > 9 && Enumerable.Range(10, lastColumn - 9).Any(c => Value(sheet.Cell(r, c)).Length > 0))
                    Issue("Colunas adicionais preservadas em Origem; confirme o seu significado.");
                var signature = Signature(name, contact, company, role, duration, activity, salary, observation, invalidSalary);
                if (!signatures.TryGetValue(signature, out var placementId))
                {
                    var companyId = "";
                    if (company.Length > 0)
                    {
                        if (!companyNames.TryGetValue(Text.Key(company), out companyId))
                        { companyId = Text.Id(); data.Companies.Add(new(companyId, company, observation)); companyNames.Add(Text.Key(company), companyId); }
                        else
                        {
                            var index = data.Companies.FindIndex(c => c.Id == companyId); var current = data.Companies[index];
                            var observations = current.Observation.Split('\n', StringSplitOptions.RemoveEmptyEntries).ToList();
                            if (observation.Length > 0 && !observations.Any(o => Text.Key(o) == Text.Key(observation)))
                            {
                                if (observations.Count > 0) Issue($"Observações diferentes para a empresa «{company}». Mantidas para revisão.");
                                observations.Add(observation); data.Companies[index] = current with { Observation = string.Join("\n", observations) };
                            }
                        }
                    }
                    string LookupId(string value, List<Lookup> list, Dictionary<string, string> index)
                    {
                        if (value.Length == 0) return "";
                        if (index.TryGetValue(Text.Key(value), out var existing)) return existing;
                        var id = Text.Id(); list.Add(new(id, value, "")); index.Add(Text.Key(value), id); return id;
                    }
                    var candidateKey = Text.Key(name) + "\n" + Text.Key(contact);
                    if (contact.Length == 0 || !contactCandidates.TryGetValue(candidateKey, out var candidateId))
                    {
                        candidateId = Text.Id(); data.Candidates.Add(new(candidateId, name, contact));
                        if (contact.Length > 0) contactCandidates[candidateKey] = candidateId;
                    }
                    placementId = Text.Id(); data.Placements.Add(new(placementId, candidateId, companyId,
                        LookupId(role, data.Roles, roleNames), LookupId(duration, data.Durations, durationNames),
                        LookupId(activity, data.Activities, activityNames), salary, observation));
                    signatures.Add(signature, placementId); imported++;
                }
                else duplicates++;
                data.Sources.Add(new(Text.Id(), placementId, System.IO.Path.GetFileName(path), sheet.Name, r, raw)); provenance.Add(sourceKey);
            }
        }
        if (!found) throw new InvalidDataException("Não foi encontrado o cabeçalho «Nome do Candidato» nas primeiras 20 linhas das folhas.");
        if (imported == 0 && duplicates == 0 && unassigned == 0) throw new InvalidDataException("O ficheiro não contém candidatos para importar.");
        ExcelStore.Validate(data);
        return new(imported, duplicates, data.Issues.Count - issuesBefore, unassigned);
    }
    private static string Value(IXLCell cell) => (cell.HasFormula ? cell.CachedValue : cell.Value).ToString(CultureInfo.InvariantCulture);
}
