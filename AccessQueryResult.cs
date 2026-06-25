using System.Text;

namespace ElectricityProductionApp;

public class AccessQueryResult
{
    public AccessQueryResult(string title, IReadOnlyList<string> columns, IReadOnlyList<IReadOnlyList<string>> rows)
    {
        Title = title;
        Columns = columns;
        Rows = rows;
    }

    public string Title { get; }
    public IReadOnlyList<string> Columns { get; }
    public IReadOnlyList<IReadOnlyList<string>> Rows { get; }

    public string ToDisplayText()
    {
        var sb = new StringBuilder();
        sb.AppendLine(Title);
        sb.AppendLine(string.Join("; ", Columns));

        foreach (IReadOnlyList<string> row in Rows)
        {
            sb.AppendLine(string.Join("; ", row));
        }

        if (Rows.Count == 0)
        {
            sb.AppendLine("Нет записей.");
        }

        return sb.ToString();
    }
}
