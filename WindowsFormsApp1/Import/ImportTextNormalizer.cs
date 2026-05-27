public static class ImportTextNormalizer
{
    public static string ToPlainText(string text)
    {
        if (string.IsNullOrEmpty(text))
            return text;

        return text
            .Replace('₀', '0')
            .Replace('₁', '1')
            .Replace('₂', '2')
            .Replace('₃', '3')
            .Replace('₄', '4')
            .Replace('₅', '5')
            .Replace('₆', '6')
            .Replace('₇', '7')
            .Replace('₈', '8')
            .Replace('₉', '9')
            .Replace('⁰', '0')
            .Replace('¹', '1')
            .Replace('²', '2')
            .Replace('³', '3')
            .Replace('⁴', '4')
            .Replace('⁵', '5')
            .Replace('⁶', '6')
            .Replace('⁷', '7')
            .Replace('⁸', '8')
            .Replace('⁹', '9');
    }
}
