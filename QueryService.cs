namespace ElectricityProductionApp;

public static class QueryService
{
    public static ElectricityRecord? CountryWithMax2015(IEnumerable<ElectricityRecord> records)
    {
        return records.OrderByDescending(record => record.Year2015).FirstOrDefault();
    }

    public static List<ElectricityRecord> CountriesWith2015MoreThan70(IEnumerable<ElectricityRecord> records)
    {
        return records
            .Where(record => record.Year2015 > 70)
            .OrderBy(record => record.Country)
            .Select(record => record.Clone())
            .ToList();
    }

    public static List<ElectricityRecord> CountriesWith2010NotMoreThan100(IEnumerable<ElectricityRecord> records)
    {
        return records
            .Where(record => record.Year2010 <= 100)
            .OrderBy(record => record.Country)
            .Select(record => record.Clone())
            .ToList();
    }

    public static string BuildSummary(IEnumerable<ElectricityRecord> records)
    {
        var list = records.ToList();
        if (list.Count == 0) return "Нет данных для расчета.";

        decimal total2010 = list.Sum(record => record.Year2010);
        decimal total2015 = list.Sum(record => record.Year2015);
        decimal avg2010 = Math.Round(list.Average(record => record.Year2010), 2);
        decimal avg2015 = Math.Round(list.Average(record => record.Year2015), 2);
        decimal totalIncrease = total2015 - total2010;
        decimal totalIncreasePercent = total2010 == 0 ? 0 : Math.Round(totalIncrease / total2010 * 100, 2);

        ElectricityRecord? max2015 = CountryWithMax2015(list);

        return "Итоговые показатели по файлу CSV" + Environment.NewLine +
               $"Всего за 2010 г.: {total2010:0.##} млрд кВт/час" + Environment.NewLine +
               $"Всего за 2015 г.: {total2015:0.##} млрд кВт/час" + Environment.NewLine +
               $"Среднее за 2010 г.: {avg2010:0.##} млрд кВт/час" + Environment.NewLine +
               $"Среднее за 2015 г.: {avg2015:0.##} млрд кВт/час" + Environment.NewLine +
               $"Абсолютный прирост: {totalIncrease:0.##} млрд кВт/час" + Environment.NewLine +
               $"Темп прироста: {totalIncreasePercent:0.##}%" + Environment.NewLine +
               $"Максимум за 2015 г.: {max2015?.Country} - {max2015?.Year2015:0.##} млрд кВт/час";
    }
}
