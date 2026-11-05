using System.Text;

namespace QuickExcel.Core;

public static class DelimitedText
{
    public static IEnumerable<string[]> ReadRows(string text, char separator, int maximumFields)
    {
        var fields = new List<string>();
        var value = new StringBuilder();
        var quoted = false;
        var closed = false;
        var count = 0;
        void Field()
        {
            if (++count > maximumFields) throw new ArgumentException(MessageText.Format($"数据上限为 {maximumFields:N0} 格。"));
            fields.Add(value.ToString()); value.Clear(); closed = false;
        }
        for (var i = 0; i < text.Length; i++)
        {
            var ch = text[i];
            if (quoted)
            {
                if (ch == '"')
                {
                    if (i + 1 < text.Length && text[i + 1] == '"') { value.Append('"'); i++; }
                    else { quoted = false; closed = true; }
                }
                else value.Append(ch);
            }
            else if (ch == '"' && value.Length == 0 && !closed) quoted = true;
            else if (ch == separator || ch is '\r' or '\n')
            {
                Field();
                if (ch != separator)
                {
                    if (ch == '\r' && i + 1 < text.Length && text[i + 1] == '\n') i++;
                    yield return fields.ToArray(); fields.Clear();
                    if (i == text.Length - 1) yield break;
                }
            }
            else
            {
                if (closed) throw new ArgumentException(MessageText.Get("字段的引号格式无效。"));
                value.Append(ch);
            }
        }
        if (quoted) throw new ArgumentException(MessageText.Get("字段的引号未闭合。"));
        Field(); yield return fields.ToArray();
    }
}
