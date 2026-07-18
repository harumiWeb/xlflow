using System.Text;
using System.Text.RegularExpressions;

namespace Xlflow.ExcelBridge.Services;

// This transformer deliberately preserves every source line and its contents.
// Erl values are therefore physical line numbers in the tracked source file.
internal static class ErlLineNumberTransformer
{
    private static readonly Regex NumberPrefix = new(@"^\s*(\d+)\s+(.+)$", RegexOptions.Compiled);
    private static readonly Regex GeneratedPrefix = new(@"^\s*\d+ (.*)$", RegexOptions.Compiled);
    private static readonly Regex ExportedGeneratedPrefix = new(@"^\s*\d+  (.*)$", RegexOptions.Compiled);
    private static readonly Regex NumericTarget = new(@"\b(?:GO\s*TO|GOSUB|RESUME)\s+(\d+)\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex ProcedureStart = new(@"^\s*(?:(?:PUBLIC|PRIVATE|FRIEND|STATIC)\s+)*(?:SUB|FUNCTION)\b|^\s*(?:(?:PUBLIC|PRIVATE|FRIEND|STATIC)\s+)*PROPERTY\s+(?:GET|LET|SET)\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex ProcedureEnd = new(@"^\s*END\s+(?:SUB|FUNCTION|PROPERTY)\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex Declaration = new(@"^\s*(?:DIM|STATIC|CONST|PRIVATE|PUBLIC|FRIEND|GLOBAL|DECLARE|TYPE|ENUM|EVENT)\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    internal sealed record SafetyIssue(int Line, string Message);

    internal static bool TryAdd(string text, out string transformed, out SafetyIssue? issue)
    {
        var lines = SplitLines(text, out var newline);
        if (!TryValidateNoExistingNumbers(lines, out issue))
        {
            transformed = text;
            return false;
        }

        var width = Math.Max(1, lines.Count.ToString(System.Globalization.CultureInfo.InvariantCulture).Length);
        var inProcedure = false;
        var continuationTail = false;
        for (var index = 0; index < lines.Count; index++)
        {
            var content = lines[index];
            if (ProcedureStart.IsMatch(content))
            {
                inProcedure = true;
                continuationTail = HasContinuation(content);
                continue;
            }
            if (ProcedureEnd.IsMatch(content))
            {
                inProcedure = false;
                continuationTail = false;
                continue;
            }

            if (inProcedure && !continuationTail && IsEligible(content))
            {
                var number = (index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture).PadLeft(width, ' ');
                lines[index] = number + " " + content;
            }
            continuationTail = HasContinuation(content);
        }

        transformed = string.Join(newline, lines);
        return true;
    }

    internal static bool TryRemove(string text, out string transformed, out SafetyIssue? issue, bool excelExported = false)
    {
        var lines = SplitLines(text, out var newline);
        if (HasNumericTarget(lines, out issue))
        {
            transformed = text;
            return false;
        }

        var hasNumbers = lines.Any(line => NumberPrefix.IsMatch(line));
        if (!hasNumbers)
        {
            transformed = text;
            return true;
        }

        var inProcedure = false;
        var continuationTail = false;
        for (var index = 0; index < lines.Count; index++)
        {
            var original = lines[index];
            var directive = NumberPrefix.Match(original);
            var generatedPrefix = excelExported ? ExportedGeneratedPrefix : GeneratedPrefix;
            var content = directive.Success ? directive.Groups[2].Value : original;
            if (ProcedureStart.IsMatch(content))
            {
                if (directive.Success)
                {
                    issue = new SafetyIssue(index + 1, "numeric line number on a procedure declaration is not xlflow-generated");
                    transformed = text;
                    return false;
                }
                inProcedure = true;
                continuationTail = HasContinuation(content);
                continue;
            }
            if (ProcedureEnd.IsMatch(content))
            {
                if (directive.Success)
                {
                    issue = new SafetyIssue(index + 1, "numeric line number on a procedure boundary is not xlflow-generated");
                    transformed = text;
                    return false;
                }
                inProcedure = false;
                continuationTail = false;
                continue;
            }

            var eligible = inProcedure && !continuationTail && IsEligible(content);
            if (directive.Success)
            {
                if (!eligible || !int.TryParse(directive.Groups[1].Value, out var number) || number != index + 1 || !generatedPrefix.IsMatch(original))
                {
                    issue = new SafetyIssue(index + 1, "numeric line label is not an xlflow-generated physical line number");
                    transformed = text;
                    return false;
                }
                lines[index] = generatedPrefix.Match(original).Groups[1].Value;
            }
            else if (eligible)
            {
                issue = new SafetyIssue(index + 1, "missing xlflow-generated line number in an executable procedure statement");
                transformed = text;
                return false;
            }
            continuationTail = HasContinuation(content);
        }

        transformed = string.Join(newline, lines);
        return true;
    }

    private static bool TryValidateNoExistingNumbers(List<string> lines, out SafetyIssue? issue)
    {
        if (HasNumericTarget(lines, out issue))
        {
            return false;
        }
        for (var index = 0; index < lines.Count; index++)
        {
            if (NumberPrefix.IsMatch(lines[index]))
            {
                issue = new SafetyIssue(index + 1, "existing numeric line label would be changed by Erl instrumentation");
                return false;
            }
        }
        issue = null;
        return true;
    }

    private static bool HasNumericTarget(List<string> lines, out SafetyIssue? issue)
    {
        for (var index = 0; index < lines.Count; index++)
        {
            var matches = NumericTarget.Matches(StripStringsAndComment(lines[index]));
            if (matches.Any(match => int.TryParse(match.Groups[1].Value, out var target) && target > 0))
            {
                issue = new SafetyIssue(index + 1, "numeric GoTo, GoSub, or Resume target would be unsafe to transform");
                return true;
            }
        }
        issue = null;
        return false;
    }

    private static bool IsEligible(string line)
    {
        var trimmed = line.Trim();
        if (string.IsNullOrWhiteSpace(trimmed) || trimmed.StartsWith('\'') || trimmed.StartsWith('#') || Declaration.IsMatch(trimmed))
        {
            return false;
        }
        if (Regex.IsMatch(trimmed, @"^[A-Za-z_][A-Za-z0-9_]*:\s*$"))
        {
            return false;
        }
        if (Regex.IsMatch(trimmed, @"^(?:SELECT\s+CASE|CASE(?:\s+ELSE)?|END\s+SELECT|ELSE|ELSEIF|END\s+(?:IF|WITH)|FOR(?:\s+EACH)?|NEXT|DO(?:\s+(?:WHILE|UNTIL))?|LOOP(?:\s+(?:WHILE|UNTIL))?|WHILE|WEND|WITH)\b", RegexOptions.IgnoreCase))
        {
            return false;
        }
        return !Regex.IsMatch(StripStringsAndComment(trimmed), @"^IF\b.*\bTHEN\s*$", RegexOptions.IgnoreCase);
    }

    private static bool HasContinuation(string line)
    {
        return StripStringsAndComment(line).TrimEnd().EndsWith(" _", StringComparison.Ordinal);
    }

    private static string StripStringsAndComment(string line)
    {
        var builder = new StringBuilder();
        var inString = false;
        for (var index = 0; index < line.Length; index++)
        {
            var ch = line[index];
            if (ch == '"')
            {
                if (inString && index + 1 < line.Length && line[index + 1] == '"')
                {
                    index++;
                    continue;
                }
                inString = !inString;
                continue;
            }
            if (!inString && ch == '\'')
            {
                break;
            }
            if (!inString)
            {
                builder.Append(ch);
            }
        }
        return builder.ToString();
    }

    private static List<string> SplitLines(string text, out string newline)
    {
        newline = text.Contains("\r\n", StringComparison.Ordinal) ? "\r\n" : text.Contains('\n') ? "\n" : text.Contains('\r') ? "\r" : Environment.NewLine;
        return text.Split(["\r\n", "\n", "\r"], StringSplitOptions.None).ToList();
    }
}
