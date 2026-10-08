using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace EmigrantesPort.Core;

public static partial class Text
{
    public static string Id() => Guid.NewGuid().ToString("D");
    public static string Clean(string? value) => Whitespace().Replace(value ?? "", " ").Trim();
    public static string Key(string? value) => string.Concat(Clean(value).Normalize(NormalizationForm.FormD)
        .Where(c => CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)).ToUpperInvariant();
    public static string Duration(string value)
    {
        var clean = Clean(value); var key = Key(clean);
        var match = Regex.Match(key, @"^(\d+)\s*MESE?S?$", RegexOptions.CultureInvariant);
        if (match.Success) return $"{match.Groups[1].Value} meses";
        return key is "INDETERIMINADO" or "INDETERMINADO" ? "Indeterminado" : clean;
    }
    public static decimal? Salary(string value)
    {
        var clean = Clean(value).Replace("€", "").Replace(" ", "");
        if (clean.Length == 0) return null;
        // Source workbook uses Portuguese thousands separators (1.150) and decimal commas.
        var format = new NumberFormatInfo { NumberDecimalSeparator = ",", NumberGroupSeparator = "." };
        if (!Regex.IsMatch(clean, @"^\+?(?:\d+|\d{1,3}(?:\.\d{3})+)(?:,\d+)?$", RegexOptions.CultureInvariant) ||
            !decimal.TryParse(clean, NumberStyles.Number, format, out var salary) || salary < 0)
            throw new InvalidDataException($"Salário inválido: {value}");
        return salary;
    }
    [GeneratedRegex(@"\s+")] private static partial Regex Whitespace();
}
