using ClosedXML.Excel;
using EmigrantesPort.Core;

var checks = 0;
void Check(bool condition, string message) { if (!condition) throw new Exception("FAIL: " + message); Console.WriteLine("PASS: " + message); checks++; }
void Throws<T>(Action action, string message) where T : Exception
{ try { action(); } catch (T) { Check(true, message); return; } throw new Exception("FAIL: " + message); }
var temp = Path.Combine(Path.GetTempPath(), "emigrantes-checks-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(temp);
try
{
    var data = new WorkbookData();
    var company = data.AddCatalog(Catalog.Companies, "Empresa Lisboa", "Fund. 2025");
    var role = data.AddCatalog(Catalog.Roles, "Electricista", "Instalações eléctricas");
    var activity = data.AddCatalog(Catalog.Activities, "Construção civil", "");
    var duration = data.AddCatalog(Catalog.Durations, "12 meses", "");
    Throws<InvalidDataException>(() => data.AddCatalog(Catalog.Roles, "  ELECTRICISTA ", ""), "Nome de função único");
    Throws<InvalidDataException>(() => data.AddCatalog(Catalog.Activities, "CONSTRUCAO CIVIL", ""), "Unicidade ignora acentos e espaços");
    var id = data.SavePlacement(new("Ana Silva", "0012345", company, role, duration, activity, 1150.50m));
    var candidateId = data.Placements.Single().CandidateId;
    var second = data.SavePlacement(new("Ana Silva", "0012345", company, role, duration, activity, 1200, ExistingCandidateId: candidateId));
    Check(data.Candidates.Count == 1 && data.Placements.Count == 2, "Uma pessoa pode ter várias contratações sem duplicação");
    data.DeletePlacement(second); Check(data.Candidates.Count == 1, "Remover um contrato preserva candidato com outra ligação");
    var path = Path.Combine(temp, "base.xlsx"); var store = new ExcelStore(path); store.Save(data);
    var loaded = new ExcelStore(path).Load();
    Check(loaded.Query().Single().Salary == 1150.50m && loaded.Query().Single().Contact == "0012345", "Gravação e leitura preservam salário e zeros do contacto");
    using (var book = new XLWorkbook(path))
    {
        Check(book.Worksheet("Candidatos").Cell(2, 1).DataType == XLDataType.Text, "IDs armazenados como texto no Excel");
        Check(book.Worksheet("Consulta").Cell(2, 3).GetString() == "Empresa Lisboa", "Consulta reúne os nomes das tabelas ligadas");
    }
    Throws<InvalidDataException>(() => data.DeleteCatalog(Catalog.Companies, company), "Remoção de empresa ligada é bloqueada");
    Throws<InvalidDataException>(() => data.SavePlacement(new("", "", company, role, duration, activity, 100)), "Candidato sem nome é rejeitado");
    Throws<InvalidDataException>(() => data.SavePlacement(new("Ana", "", "bad", role, duration, activity, 100)), "Referências inexistentes são rejeitadas");
    Throws<InvalidDataException>(() => data.SavePlacement(new("Ana", "", company, role, duration, activity, -1)), "Salário negativo é rejeitado");
    data.EditCatalog(Catalog.Companies, company, "Empresa Porto", "Fund. 2025"); store.Save(data);
    Check(new ExcelStore(path).Load().Query().Single().Company == "Empresa Porto", "Alteração do catálogo actualiza a consulta");
    Check(Directory.GetFiles(Path.Combine(temp, "CopiasSeguranca")).Length == 1, "Gravação cria uma cópia de segurança");
    Check(data.Query("ana", company, role).Count == 1 && data.Query("inexistente").Count == 0, "Pesquisa e filtros funcionam");
    Check(Text.Salary("1.150") == 1150m && Text.Salary("1.150,50 €") == 1150.50m, "Salários portugueses são interpretados correctamente");
    Throws<InvalidDataException>(() => Text.Salary("1.5"), "Salário com separadores ambíguos é rejeitado");
    Check(Text.Duration("12meses") == "12 meses" && Text.Duration("indeteriminado") == "Indeterminado", "Durações conhecidas são normalizadas");
    var other = new ExcelStore(path); var edited = other.Load(); edited.EditCatalog(Catalog.Roles, role, "Instalador", ""); other.Save(edited);
    Throws<IOException>(() => store.Save(data), "Alteração externa impede sobrescrita silenciosa");
    data.DeletePlacement(id); Check(data.Placements.Count == 0 && data.Candidates.Count == 0, "Remoção limpa candidato sem contratações");
    data.DeleteCatalog(Catalog.Companies, company); Check(data.Companies.Count == 0, "Empresa sem ligações pode ser removida");
    var invalid = new WorkbookData(); invalid.Candidates.Add(new("x", "A", "")); invalid.Placements.Add(new("p", "x", "broken", "", "", "", null, ""));
    Throws<InvalidDataException>(() => ExcelStore.Validate(invalid), "Integridade referencial é validada na base");
    var fixture = Path.Combine(temp, "original.xlsx");
    using (var book = new XLWorkbook())
    {
        var sheet = book.AddWorksheet("Folha1"); var headers = new[] { "Ord.", "Nome do Candidato", "Empresa Contratante", "Função a exercer", "Duração do ", "Actividade", "Salário", "Contacto", "Observação" };
        for (var c = 0; c < headers.Length; c++) sheet.Cell(1, c + 1).Value = headers[c];
        var values = new[] { "1", "Ana Silva", "Empresa A", "Electricista", "12meses", "Construção civil", "1.150", "00123", "Fund. 2025" };
        for (var r = 2; r <= 3; r++) for (var c = 0; c < values.Length; c++) sheet.Cell(r, c + 1).Value = values[c];
        sheet.Cell(4, 2).Value = "Bento"; sheet.Cell(4, 3).Value = "Empresa A"; sheet.Cell(4, 7).Value = "por confirmar"; sheet.Cell(4, 9).Value = "Fund. 2020";
        sheet.Cell(4, 10).Value = "nota adicional";
        book.SaveAs(fixture);
    }
    var imported = new WorkbookData(); var report = ExcelImporter.Import(imported, fixture);
    Check(report.Imported == 2 && report.Duplicates == 1 && imported.Sources.Count == 3, "Importação elimina linhas idênticas e preserva todas as origens");
    Check(imported.Companies.Count == 1 && imported.Companies[0].Observation.Contains("2020") && imported.Companies[0].Observation.Contains("2025"), "Observações divergentes são preservadas");
    Check(imported.Query().Any(p => p.Salary is null) && imported.Issues.Any(i => i.Message.Contains("adicionais")), "Valores incompletos e colunas extras ficam para revisão");
    var twice = ExcelImporter.Import(imported, fixture);
    Check(twice.Imported == 0 && imported.Sources.Count == 3 && twice.Issues == 0, "Importação repetida é idempotente");
    imported.DeletePlacement(imported.Placements.First(p => imported.Sources.Count(s => s.PlacementId == p.Id) == 2).Id);
    Check(imported.Sources.Count == 1, "Remover contratação limpa todas as suas linhas de origem");
    var normalized = Path.Combine(temp, "normalizada.xlsx"); var normalizedStore = new ExcelStore(normalized); normalizedStore.Save(imported);
    Check(normalizedStore.Load().Sources.Count == 1, "Origem e tabelas sobrevivem à leitura da base normalizada");
    Throws<InvalidDataException>(() => ExcelImporter.Import(new WorkbookData(), normalized), "Base normalizada deve ser aberta, não importada como original");
    var formulaData = new WorkbookData(); var fcompany = formulaData.AddCatalog(Catalog.Companies, "=TESTE", ""); var formulaFile = Path.Combine(temp, "formula.xlsx"); new ExcelStore(formulaFile).Save(formulaData);
    using (var book = new XLWorkbook(formulaFile)) Check(!book.Worksheet("Empresas").Cell(2, 2).HasFormula, "Textos iniciados por = não se tornam fórmulas");
    SessionChecks.Run(temp, Check);
    if (args.Length > 0)
    {
        var originalHash = System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(args[0]));
        var real = new WorkbookData(); var result = ExcelImporter.Import(real, args[0]);
        Console.WriteLine($"REAL WORKBOOK: {result.Imported} contratações; {result.Duplicates} duplicados; {real.Candidates.Count} candidatos; {real.Companies.Count} empresas; {real.Roles.Count} funções; {real.Sources.Count} origens; {result.Issues} avisos");
        Check(real.Sources.Count == 3526, "Todas as 3526 linhas com dados do ficheiro fornecido foram preservadas");
        Check(result.Imported + result.Duplicates + result.Unassigned == 3526, "Todos os registos do ficheiro têm um resultado de importação");
        var realPath = Path.Combine(temp, "real-normalized.xlsx"); var realStore = new ExcelStore(realPath); realStore.Save(real);
        var realLoaded = realStore.Load(); Check(realLoaded.Query().Count == real.Placements.Count && realLoaded.Sources.Count == 3526, "Base real normalizada é gravada e relida sem perda de registos");
        var repeat = ExcelImporter.Import(realLoaded, args[0]); Check(repeat.Imported == 0 && repeat.Issues == 0, "Reimportar o documento real não cria duplicados");
        Check(originalHash.SequenceEqual(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(args[0]))), "Excel fornecido permanece intacto");
    }
    Console.WriteLine($"{checks} checks passed.");
}
finally { Directory.Delete(temp, true); }
