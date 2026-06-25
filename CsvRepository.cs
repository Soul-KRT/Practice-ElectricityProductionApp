using System.Globalization;
using System.Text;

namespace ElectricityProductionApp;

public class CsvRepository
{
    private readonly string _filePath;
    private static readonly CultureInfo Culture = CultureInfo.InvariantCulture;

    public CsvRepository(string filePath)
    {
        _filePath = filePath;
    }

    public List<ElectricityRecord> Load()
    {
        if (!File.Exists(_filePath))
        {
            Save(CreateDefaultRecords());
        }

        var records = new List<ElectricityRecord>();
        foreach (string line in File.ReadAllLines(_filePath, Encoding.UTF8).Skip(1))
        {
            if (string.IsNullOrWhiteSpace(line)) continue;
            string[] parts = line.Split(';');
            if (parts.Length < 4) continue;

            records.Add(new ElectricityRecord
            {
                Region = parts[0].Trim(),
                Country = parts[1].Trim(),
                Year2010 = decimal.Parse(parts[2].Replace(',', '.'), Culture),
                Year2015 = decimal.Parse(parts[3].Replace(',', '.'), Culture)
            });
        }

        return records;
    }

    public void Save(IEnumerable<ElectricityRecord> records)
    {
        string? directory = Path.GetDirectoryName(_filePath);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);

        var lines = new List<string> { "Регион;Страна;2010;2015" };
        lines.AddRange(records.Select(record => string.Join(';',
            record.Region,
            record.Country,
            ToCsvNumber(record.Year2010),
            ToCsvNumber(record.Year2015))));

        File.WriteAllLines(_filePath, lines, Encoding.UTF8);
    }

    public static List<ElectricityRecord> CreateDefaultRecords()
    {
        return new List<ElectricityRecord>
        {
            new() { Region = "Северная Америка", Country = "США", Year2010 = 629.0m, Year2015 = 724.0m },
            new() { Region = "Европа", Country = "Англия", Year2010 = 93.9m, Year2015 = 112.9m },
            new() { Region = "Северная Америка", Country = "Канада", Year2010 = 82.8m, Year2015 = 96.7m },
            new() { Region = "Европа", Country = "Германия", Year2010 = 76.5m, Year2015 = 95.1m },
            new() { Region = "Азия", Country = "Япония", Year2010 = 65.2m, Year2015 = 81.2m },
            new() { Region = "Европа", Country = "Франция", Year2010 = 49.6m, Year2015 = 61.8m },
            new() { Region = "Европа", Country = "Швеция", Year2010 = 24.7m, Year2015 = 32.5m }
        };
    }

    private static string ToCsvNumber(decimal value)
    {
        return value.ToString("0.##", Culture);
    }
}
