namespace EmigrantesPort;

public sealed record Choice(string Id, string Name)
{
    public override string ToString() => Name;
}
