namespace EmigrantesPort;

public interface IWorkbookDestination
{
    Task<string?> SaveCopyAsync(byte[] workbook, string suggestedName, string currentPath);
}
