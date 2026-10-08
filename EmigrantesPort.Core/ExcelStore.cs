using ClosedXML.Excel;
using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;

namespace EmigrantesPort.Core;

/// <summary>Excel is the only persistent database; Consulta is regenerated on each save.</summary>
public sealed class ExcelStore
{
    public string Path { get; }
    private string? fingerprint;
    public ExcelStore(string path) => Path = System.IO.Path.GetFullPath(path);
    private static string Hash(string path)
    {
        using var file = File.OpenRead(path); return Convert.ToHexString(SHA256.HashData(file));
    }
    public WorkbookData Load()
    {
        if (!File.Exists(Path)) throw new FileNotFoundException("A base Excel não foi encontrada.", Path);
        var before = Hash(Path);
        using var book = new XLWorkbook(Path);
        if (!book.TryGetWorksheet("Configuracao", out var config) || config.Cell(1, 1).GetString() != "EmigrantesPort" || config.Cell(1, 2).GetString() != "1")
            throw new InvalidDataException("Este ficheiro não é uma base normalizada. Utilize «Importar Excel» para importar a folha original.");
        var data = new WorkbookData();
        foreach (var r in Rows(book, "Candidatos", ["Id", "Nome", "Contacto"])) data.Candidates.Add(new(S(r, 1), S(r, 2), S(r, 3)));
        foreach (var r in Rows(book, "Empresas", ["Id", "Nome", "Observacao"])) data.Companies.Add(new(S(r, 1), S(r, 2), S(r, 3)));
        ReadLookups(book, "Funcoes", data.Roles); ReadLookups(book, "Actividades", data.Activities); ReadLookups(book, "Duracoes", data.Durations);
        foreach (var r in Rows(book, "Contratacoes", ["Id", "CandidatoId", "EmpresaId", "FuncaoId", "DuracaoId", "ActividadeId", "Salario", "ObservacaoOriginal"]))
            data.Placements.Add(new(S(r, 1), S(r, 2), S(r, 3), S(r, 4), S(r, 5), S(r, 6), r.Cell(7).IsEmpty() ? null : r.Cell(7).GetValue<decimal>(), S(r, 8)));
        foreach (var r in Rows(book, "Origem", ["Id", "ContratacaoId", "Ficheiro", "Folha", "Linha", "ValoresOriginais"]))
            data.Sources.Add(new(S(r, 1), S(r, 2), S(r, 3), S(r, 4), r.Cell(5).GetValue<int>(), S(r, 6)));
        foreach (var r in Rows(book, "Revisao", ["Id", "Origem", "Mensagem"])) data.Issues.Add(new(S(r, 1), S(r, 2), S(r, 3)));
        Validate(data);
        if (Hash(Path) != before) throw new IOException("O ficheiro mudou durante a leitura. Volte a abrir a base.");
        fingerprint = before; return data;
    }
    public void Save(WorkbookData data)
    {
        Validate(data);
        var directory = System.IO.Path.GetDirectoryName(Path)!; Directory.CreateDirectory(directory);
        if (File.Exists(Path) && (fingerprint is null || Hash(Path) != fingerprint))
            throw new IOException("A base foi alterada fora da aplicação. Reabra-a antes de guardar, para não perder alterações.");
        var temp = System.IO.Path.Combine(directory, $".{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var book = Build(data)) book.SaveAs(temp);
            // Check again after workbook generation; a locked Excel file leaves the current database intact.
            if (File.Exists(Path))
            {
                if (Hash(Path) != fingerprint) throw new IOException("A base foi alterada no Excel. Reabra-a antes de guardar.");
                var backup = System.IO.Path.Combine(directory, "CopiasSeguranca"); Directory.CreateDirectory(backup);
                File.Copy(Path, System.IO.Path.Combine(backup, $"{System.IO.Path.GetFileNameWithoutExtension(Path)}-{DateTime.UtcNow:yyyyMMdd-HHmmss}-{Guid.NewGuid():N}.xlsx"));
                File.Replace(temp, Path, null);
            }
            else
            {
                if (fingerprint is not null) throw new IOException("A base foi removida fora da aplicação. Reabra ou crie outra base.");
                File.Move(temp, Path);
            }
            fingerprint = Hash(Path);
        }
        finally { if (File.Exists(temp)) File.Delete(temp); }
    }
    public static void Export(WorkbookData data, string destination)
    {
        if (File.Exists(destination)) throw new IOException("Já existe um ficheiro nesse destino. Escolha outro nome.");
        new ExcelStore(destination).Save(data);
    }
    private static void ReadLookups(XLWorkbook book, string sheet, List<Lookup> list)
    { foreach (var r in Rows(book, sheet, ["Id", "Nome", "Descricao"])) list.Add(new(S(r, 1), S(r, 2), S(r, 3))); }
    private static IEnumerable<IXLRow> Rows(XLWorkbook book, string name, string[] headers)
    {
        if (!book.TryGetWorksheet(name, out var sheet)) throw new InvalidDataException($"Falta a folha {name}.");
        for (var c = 0; c < headers.Length; c++)
            if (sheet.Cell(1, c + 1).GetString() != headers[c]) throw new InvalidDataException($"Cabeçalhos inválidos na folha {name}.");
        var last = sheet.LastRowUsed(XLCellsUsedOptions.Contents)?.RowNumber() ?? 1;
        if (last > 100001) throw new InvalidDataException("A base excede o limite de 100 000 linhas por folha.");
        for (var i = 2; i <= last; i++)
        {
            var row = sheet.Row(i);
            if (row.Cells(1, headers.Length).All(c => c.IsEmpty())) continue;
            if (row.Cells(1, headers.Length).Any(c => c.HasFormula)) throw new InvalidDataException($"Fórmulas não são permitidas nas tabelas da base: {name}, linha {i}.");
            yield return row;
        }
    }
    private static string S(IXLRow row, int column) => row.Cell(column).GetString();
    public static void Validate(WorkbookData data)
    {
        static void Ids(IEnumerable<string> ids, string name)
        {
            var values = ids.ToList(); if (values.Any(string.IsNullOrWhiteSpace) || values.Distinct(StringComparer.Ordinal).Count() != values.Count)
                throw new InvalidDataException($"IDs vazios ou repetidos em {name}.");
        }
        static void Unique(IEnumerable<string> names, string table)
        {
            var keys = names.Select(Text.Key).ToList();
            if (keys.Any(string.IsNullOrEmpty) || keys.Distinct().Count() != keys.Count) throw new InvalidDataException($"Os nomes em {table} devem ser preenchidos e únicos.");
        }
        Ids(data.Candidates.Select(x => x.Id), "Candidatos"); Ids(data.Companies.Select(x => x.Id), "Empresas");
        foreach (var list in new[] { data.Roles, data.Activities, data.Durations }) { Ids(list.Select(x => x.Id), "Catálogos"); Unique(list.Select(x => x.Name), "Catálogos"); }
        Ids(data.Placements.Select(x => x.Id), "Contratações"); Ids(data.Sources.Select(x => x.Id), "Origem"); Ids(data.Issues.Select(x => x.Id), "Revisão");
        Unique(data.Companies.Select(x => x.Name), "Empresas");
        if (data.Candidates.Any(c => string.IsNullOrWhiteSpace(c.Name))) throw new InvalidDataException("Existem candidatos sem nome.");
        var candidates = data.Candidates.Select(x => x.Id).ToHashSet(); var companies = data.Companies.Select(x => x.Id).ToHashSet();
        var roles = data.Roles.Select(x => x.Id).ToHashSet(); var durations = data.Durations.Select(x => x.Id).ToHashSet(); var activities = data.Activities.Select(x => x.Id).ToHashSet();
        foreach (var p in data.Placements)
            if (!candidates.Contains(p.CandidateId) || (p.CompanyId.Length > 0 && !companies.Contains(p.CompanyId)) ||
                (p.RoleId.Length > 0 && !roles.Contains(p.RoleId)) || (p.DurationId.Length > 0 && !durations.Contains(p.DurationId)) ||
                (p.ActivityId.Length > 0 && !activities.Contains(p.ActivityId)) || p.Salary < 0)
                throw new InvalidDataException($"Ligações ou salário inválidos na contratação {p.Id}.");
        var placements = data.Placements.Select(p => p.Id).ToHashSet();
        if (data.Sources.Any(s => s.PlacementId.Length > 0 && !placements.Contains(s.PlacementId))) throw new InvalidDataException("A folha Origem contém uma contratação inexistente.");
    }
    public static XLWorkbook Build(WorkbookData data)
    {
        Validate(data); var book = new XLWorkbook();
        var config = book.AddWorksheet("Configuracao"); config.Cell(1, 1).Value = "EmigrantesPort"; config.Cell(1, 2).Value = "1";
        config.Cell(2, 1).Value = "Consulta é uma vista gerada. Edite apenas as tabelas; mantenha os IDs e as ligações.";
        Write(book, "Candidatos", ["Id", "Nome", "Contacto"], data.Candidates.Select(c => new object?[] { c.Id, c.Name, c.Contact }));
        Write(book, "Empresas", ["Id", "Nome", "Observacao"], data.Companies.Select(c => new object?[] { c.Id, c.Name, c.Observation }));
        WriteLookup(book, "Funcoes", data.Roles); WriteLookup(book, "Actividades", data.Activities); WriteLookup(book, "Duracoes", data.Durations);
        Write(book, "Contratacoes", ["Id", "CandidatoId", "EmpresaId", "FuncaoId", "DuracaoId", "ActividadeId", "Salario", "ObservacaoOriginal"],
            data.Placements.Select(p => new object?[] { p.Id, p.CandidateId, p.CompanyId, p.RoleId, p.DurationId, p.ActivityId, p.Salary, p.Observation }));
        var companies = data.Companies.ToDictionary(c => c.Id);
        var placements = data.Placements.ToDictionary(p => p.Id);
        Write(book, "Consulta", ["ContratacaoId", "Nome do Candidato", "Empresa Contratante", "Função a exercer", "Duração do Contrato", "Actividade", "Salário", "Contacto", "Observação da empresa", "Observação original"],
            data.Query().Select(p => new object?[] { p.Id, p.Name, p.Company, p.Role, p.Duration, p.Activity, p.Salary, p.Contact,
                companies.GetValueOrDefault(placements[p.Id].CompanyId)?.Observation ?? "", p.Observation }));
        Write(book, "Origem", ["Id", "ContratacaoId", "Ficheiro", "Folha", "Linha", "ValoresOriginais"], data.Sources.Select(s => new object?[] { s.Id, s.PlacementId, s.File, s.Sheet, s.Row, s.OriginalJson }));
        Write(book, "Revisao", ["Id", "Origem", "Mensagem"], data.Issues.Select(i => new object?[] { i.Id, i.Source, i.Message }));
        return book;
    }
    private static void WriteLookup(XLWorkbook book, string sheet, IEnumerable<Lookup> values) => Write(book, sheet, ["Id", "Nome", "Descricao"], values.Select(x => new object?[] { x.Id, x.Name, x.Description }));
    private static void Write(XLWorkbook book, string sheet, string[] headers, IEnumerable<object?[]> rows)
    {
        var ws = book.AddWorksheet(sheet);
        for (var c = 0; c < headers.Length; c++) ws.Cell(1, c + 1).Value = headers[c];
        var r = 2;
        foreach (var row in rows)
        {
            for (var c = 0; c < row.Length; c++)
            {
                var cell = ws.Cell(r, c + 1);
                switch (row[c])
                {
                    case null: break;
                    case decimal value: cell.Value = value; cell.Style.NumberFormat.Format = "#,##0.00"; break;
                    case int number: cell.Value = number; break;
                    default: cell.Style.NumberFormat.Format = "@"; cell.Value = row[c]!.ToString()!; break;
                }
            }
            r++;
        }
        ws.Range(1, 1, Math.Max(r - 1, 1), headers.Length).SetAutoFilter(); ws.SheetView.FreezeRows(1);
        ws.Row(1).Style.Fill.BackgroundColor = XLColor.FromHtml("#173F35"); ws.Row(1).Style.Font.FontColor = XLColor.White; ws.Row(1).Style.Font.Bold = true;
        ws.Columns(1, headers.Length).Width = 26; ws.Column(1).Width = 38;
        if (sheet == "Consulta") { ws.Column(2).Width = 36; ws.Column(3).Width = 48; }
        if (sheet == "Origem") ws.Column(6).Width = 60;
    }
}
