namespace EmigrantesPort.Core;

/// <summary>Keeps the in-memory database aligned with the last successful disk operation.</summary>
public sealed class WorkbookSession
{
    private readonly object sync = new();
    private ExcelStore? store;
    private long revision;

    public WorkbookData Data { get; private set; } = new();
    public string? Path => store?.Path;
    public bool IsOpen => store is not null;

    public void Open(string path, bool createIfMissing = false)
    {
        lock (sync)
        {
            var nextStore = new ExcelStore(path);
            WorkbookData loaded;
            if (File.Exists(path))
            {
                loaded = nextStore.Load();
            }
            else if (createIfMissing)
            {
                loaded = new WorkbookData();
                nextStore.Save(loaded);
            }
            else
            {
                throw new FileNotFoundException("A base Excel não foi encontrada. Abra outra base ou restaure uma cópia de segurança.", path);
            }

            store = nextStore;
            Data = loaded;
            revision++;
        }
    }

    public void Reload()
    {
        lock (sync)
        {
            var loaded = RequireStore().Load();
            Data = loaded;
            revision++;
        }
    }

    public void Apply(Action<WorkbookData> mutation)
    {
        lock (sync)
        {
            var currentStore = RequireStore();
            var copy = Data.Copy();
            mutation(copy);
            currentStore.Save(copy);
            Data = copy;
            revision++;
        }
    }

    public PreparedImport PrepareImport(string sourcePath)
    {
        lock (sync)
        {
            RequireStore();
            var copy = Data.Copy();
            var result = ExcelImporter.Import(copy, sourcePath);
            return new PreparedImport(this, revision, copy, result);
        }
    }

    public void CommitImport(PreparedImport prepared)
    {
        lock (sync)
        {
            if (!ReferenceEquals(prepared.Owner, this) || prepared.Revision != revision)
                throw new InvalidOperationException("A base mudou desde a preparação da importação. Prepare a importação novamente.");

            RequireStore().Save(prepared.Data);
            Data = prepared.Data;
            revision++;
        }
    }

    private ExcelStore RequireStore() => store ?? throw new InvalidOperationException("Abra a base Excel antes de continuar.");
}

public sealed class PreparedImport
{
    internal WorkbookSession Owner { get; }
    internal long Revision { get; }
    internal WorkbookData Data { get; }
    public ImportResult Result { get; }

    internal PreparedImport(WorkbookSession owner, long revision, WorkbookData data, ImportResult result)
    {
        Owner = owner;
        Revision = revision;
        Data = data;
        Result = result;
    }
}
