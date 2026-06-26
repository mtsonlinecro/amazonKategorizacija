using System.Data;
using System.Text;

namespace AmazonCategoryMapperWinForms;

public static class CsvUtil
{
    public static DataTable ReadCsv(string path)
    {
        var text = File.ReadAllText(path, Encoding.UTF8);
        var records = ParseRecords(text);
        var table = new DataTable();
        if (records.Count == 0) return table;

        foreach (var header in records[0])
        {
            var name = string.IsNullOrWhiteSpace(header) ? $"Column{table.Columns.Count + 1}" : header.Trim();
            if (table.Columns.Contains(name)) name = name + "_" + table.Columns.Count;
            table.Columns.Add(name);
        }

        for (int i = 1; i < records.Count; i++)
        {
            if (records[i].Count == 1 && string.IsNullOrWhiteSpace(records[i][0])) continue;
            var row = table.NewRow();
            for (int c = 0; c < table.Columns.Count; c++)
            {
                row[c] = c < records[i].Count ? records[i][c] : "";
            }
            table.Rows.Add(row);
        }
        return table;
    }

    public static void WriteCsv(DataTable table, string path)
    {
        var sb = new StringBuilder();
        sb.AppendLine(string.Join(",", table.Columns.Cast<DataColumn>().Select(c => Escape(c.ColumnName))));
        foreach (DataRow row in table.Rows)
        {
            sb.AppendLine(string.Join(",", table.Columns.Cast<DataColumn>().Select(c => Escape(Convert.ToString(row[c]) ?? ""))));
        }
        File.WriteAllText(path, sb.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
    }

    public static List<Dictionary<string, object?>> ToRows(DataTable table)
    {
        var list = new List<Dictionary<string, object?>>();
        foreach (DataRow row in table.Rows)
        {
            var dict = new Dictionary<string, object?>();
            foreach (DataColumn col in table.Columns)
            {
                dict[col.ColumnName] = Convert.ToString(row[col]) ?? "";
            }
            list.Add(dict);
        }
        return list;
    }

    private static string Escape(string value)
    {
        value ??= "";
        if (value.Contains('"') || value.Contains(',') || value.Contains('\n') || value.Contains('\r'))
        {
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }
        return value;
    }

    private static List<List<string>> ParseRecords(string text)
    {
        var records = new List<List<string>>();
        var row = new List<string>();
        var field = new StringBuilder();
        bool inQuotes = false;

        for (int i = 0; i < text.Length; i++)
        {
            char ch = text[i];
            if (inQuotes)
            {
                if (ch == '"')
                {
                    if (i + 1 < text.Length && text[i + 1] == '"')
                    {
                        field.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = false;
                    }
                }
                else
                {
                    field.Append(ch);
                }
            }
            else
            {
                if (ch == '"')
                {
                    inQuotes = true;
                }
                else if (ch == ',')
                {
                    row.Add(field.ToString());
                    field.Clear();
                }
                else if (ch == '\r')
                {
                    // ignore, \n će završiti red
                }
                else if (ch == '\n')
                {
                    row.Add(field.ToString());
                    field.Clear();
                    records.Add(row);
                    row = new List<string>();
                }
                else
                {
                    field.Append(ch);
                }
            }
        }

        row.Add(field.ToString());
        if (row.Count > 1 || !string.IsNullOrEmpty(row[0])) records.Add(row);
        return records;
    }
}
