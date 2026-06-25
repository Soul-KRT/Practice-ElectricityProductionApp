namespace ElectricityProductionApp;

public class ElectricityRecord
{
    public string Region { get; set; } = string.Empty;
    public string Country { get; set; } = string.Empty;
    public decimal Year2010 { get; set; }
    public decimal Year2015 { get; set; }

    public decimal Increase => Year2015 - Year2010;

    public decimal IncreasePercent
    {
        get
        {
            if (Year2010 == 0) return 0;
            return Math.Round((Year2015 - Year2010) / Year2010 * 100, 2);
        }
    }

    public ElectricityRecord Clone()
    {
        return new ElectricityRecord
        {
            Region = Region,
            Country = Country,
            Year2010 = Year2010,
            Year2015 = Year2015
        };
    }
}
