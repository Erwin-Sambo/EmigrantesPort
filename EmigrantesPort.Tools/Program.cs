using EmigrantesPort.Core;

try
{
    if (args.Length == 3 && args[0] == "normalizar")
    {
        if (File.Exists(args[2])) throw new IOException("O destino já existe. Escolha outro ficheiro para preservar os dados.");
        var data = new WorkbookData(); var report = ExcelImporter.Import(data, args[1]);
        new ExcelStore(args[2]).Save(data);
        // A successful conversion also requires a readable result.
        var reopened = new ExcelStore(args[2]).Load();
        Console.WriteLine($"Base criada: {Path.GetFullPath(args[2])}");
        Console.WriteLine($"{reopened.Candidates.Count} candidatos · {reopened.Companies.Count} empresas · {reopened.Placements.Count} contratações");
        Console.WriteLine($"{report.Duplicates} linhas repetidas · {report.Unassigned} linhas sem candidato · {report.Issues} pontos para revisão");
        return 0;
    }
    if (args.Length is 2 or 3 && args[0] == "consultar")
    {
        var data = new ExcelStore(args[1]).Load(); var rows = data.Query(args.Length == 3 ? args[2] : "");
        Console.WriteLine("Candidato\tEmpresa\tFunção\tDuração\tActividade\tSalário\tContacto");
        foreach (var row in rows) Console.WriteLine(string.Join('\t', new[] { row.Name, row.Company, row.Role, row.Duration, row.Activity, row.SalaryText, row.Contact }));
        Console.WriteLine($"{rows.Count} resultados"); return 0;
    }
    Console.Error.WriteLine("Utilização: normalizar original.xlsx nova-base.xlsx | consultar base.xlsx [pesquisa]"); return 2;
}
catch (Exception ex) { Console.Error.WriteLine("Operação não concluída: " + ex.Message); return 1; }
