using ClosedXML.Excel;
using EmigrantesPort.Core;

internal static class SessionChecks
{
    internal static void Run(string directory, Action<bool, string> check)
    {
        void Reject<T>(Action action, string message) where T : Exception
        {
            try { action(); }
            catch (T) { check(true, message); return; }
            check(false, message);
        }

        var session = new WorkbookSession();
        var path = Path.Combine(directory, "session.xlsx");
        session.Open(path, createIfMissing: true);
        session.Apply(data => data.AddCatalog(Catalog.Companies, "Empresa original", ""));
        Reject<FileNotFoundException>(() => session.Open(Path.Combine(directory, "missing.xlsx")), "Abrir ficheiro em falta preserva a sessão anterior");
        check(session.Path == path && session.Data.Companies.Single().Name == "Empresa original", "Sessão mantém a base após falha ao abrir");
        Reject<InvalidDataException>(() => session.Apply(data =>
        {
            data.AddCatalog(Catalog.Roles, "Função", "");
            data.AddCatalog(Catalog.Roles, "FUNCAO", "");
        }), "Alteração inválida é rejeitada");
        check(session.Data.Roles.Count == 0 && new ExcelStore(path).Load().Roles.Count == 0, "Falha de edição não altera memória nem disco");

        var fixture = Path.Combine(directory, "session-import.xlsx");
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.AddWorksheet("Dados");
            sheet.Cell(1, 2).Value = "Nome do Candidato";
            sheet.Cell(2, 2).Value = "Candidato de teste";
            workbook.SaveAs(fixture);
        }
        var prepared = session.PrepareImport(fixture);
        check(session.Data.Candidates.Count == 0 && prepared.Result.Imported == 1, "Preparar importação não modifica a base");
        session.Apply(data => data.AddCatalog(Catalog.Roles, "Nova função", ""));
        Reject<InvalidOperationException>(() => session.CommitImport(prepared), "Importação antiga não sobrescreve alterações posteriores");
        prepared = session.PrepareImport(fixture);
        session.CommitImport(prepared);
        check(session.Data.Candidates.Count == 1 && new ExcelStore(path).Load().Candidates.Count == 1, "Importação confirmada actualiza memória e disco");
        Reject<InvalidOperationException>(() => session.CommitImport(prepared), "Uma preparação não pode ser confirmada duas vezes");

        var external = new ExcelStore(path);
        var edited = external.Load();
        edited.AddCatalog(Catalog.Companies, "Alteração externa", "");
        external.Save(edited);
        Reject<IOException>(() => session.Apply(data => data.AddCatalog(Catalog.Roles, "Não deve entrar", "")), "Conflito externo impede gravação da sessão");
        check(!session.Data.Roles.Any(r => r.Name == "Não deve entrar"), "Conflito de gravação não publica alterações em memória");
        session.Reload();
        check(session.Data.Companies.Any(c => c.Name == "Alteração externa"), "Reabrir carrega alterações externas");

        var invalidWorkbook = Path.Combine(directory, "late-failure.xlsx");
        using (var workbook = new XLWorkbook())
        {
            var first = workbook.AddWorksheet("Primeira");
            first.Cell(1, 2).Value = "Nome do Candidato";
            first.Cell(2, 2).Value = "Registo válido";
            var second = workbook.AddWorksheet("Demasiado larga");
            second.Cell(1, 2).Value = "Nome do Candidato";
            second.Cell(2, 201).Value = "Excede o limite";
            workbook.SaveAs(invalidWorkbook);
        }
        var target = new WorkbookData();
        Reject<InvalidDataException>(() => ExcelImporter.Import(target, invalidWorkbook), "Falha numa folha posterior cancela a importação");
        check(target.Candidates.Count == 0 && target.Sources.Count == 0 && target.Issues.Count == 0, "Importação falhada não deixa registos parciais");

        var salaries = Path.Combine(directory, "invalid-salaries.xlsx");
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.AddWorksheet("Dados");
            sheet.Cell(1, 2).Value = "Nome do Candidato";
            sheet.Cell(2, 2).Value = "Ana"; sheet.Cell(2, 7).Value = "valor A";
            sheet.Cell(3, 2).Value = "Ana"; sheet.Cell(3, 7).Value = "valor B";
            workbook.SaveAs(salaries);
        }
        var salaryData = new WorkbookData();
        var report = ExcelImporter.Import(salaryData, salaries);
        check(report.Imported == 2 && report.Duplicates == 0, "Salários inválidos diferentes não são tratados como duplicados");
    }
}
