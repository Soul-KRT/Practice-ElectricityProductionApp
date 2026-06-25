using System.Globalization;
using System.Runtime.InteropServices;

namespace ElectricityProductionApp;

public class AccessRepository
{
    private readonly string _dbPath;

    public AccessRepository(string dbPath)
    {
        _dbPath = dbPath;
    }

    public bool Exists()
    {
        return File.Exists(_dbPath);
    }

    public void Recreate()
    {
        AccessDatabaseCreator.Create(_dbPath, overwrite: true);
    }

    public AccessQueryResult LoadAll()
    {
        const string sql = @"
SELECT R.RegionName, EP.CountryName, EP.Production2010, EP.Production2015
FROM Regions AS R
INNER JOIN EnergyProduction AS EP ON R.RegionId = EP.RegionId
ORDER BY EP.CountryName";

        return ExecuteQuery("MS Access. Исходная таблица варианта 13", sql);
    }

    public AccessQueryResult QueryTopCountry2015()
    {
        const string sql = @"
SELECT TOP 1 EP.CountryName, EP.Production2015
FROM EnergyProduction AS EP
ORDER BY EP.Production2015 DESC";

        return ExecuteQuery("Access-запрос 1. Страна с максимумом производства в 2015 г.", sql);
    }

    public AccessQueryResult QueryCountriesOver70In2015()
    {
        const string sql = @"
SELECT R.RegionName, EP.CountryName, EP.Production2015
FROM Regions AS R
INNER JOIN EnergyProduction AS EP ON R.RegionId = EP.RegionId
WHERE EP.Production2015 > 70
ORDER BY EP.CountryName";

        return ExecuteQuery("Access-запрос 2. Производство в 2015 г. превысило 70 млрд кВт/час", sql);
    }

    public AccessQueryResult QueryCountriesNotOver100In2010()
    {
        const string sql = @"
SELECT R.RegionName, EP.CountryName, EP.Production2010
FROM Regions AS R
INNER JOIN EnergyProduction AS EP ON R.RegionId = EP.RegionId
WHERE EP.Production2010 <= 100
ORDER BY EP.CountryName";

        return ExecuteQuery("Access-запрос 3. Производство в 2010 г. не превышало 100 млрд кВт/час", sql);
    }

    public AccessQueryResult ExecuteQuery(string title, string sql)
    {
        dynamic? connection = null;
        dynamic? recordset = null;

        try
        {
            CheckDatabaseFile();
            connection = AccessDatabaseCreator.OpenConnection(_dbPath);
            recordset = connection.Execute(sql);

            var columns = new List<string>();
            int fieldCount = Convert.ToInt32(recordset.Fields.Count, CultureInfo.InvariantCulture);
            for (int i = 0; i < fieldCount; i++)
            {
                dynamic field = recordset.Fields[i];
                columns.Add(Convert.ToString(field.Name, CultureInfo.InvariantCulture) ?? string.Empty);
            }

            var rows = new List<IReadOnlyList<string>>();
            while (!Convert.ToBoolean(recordset.EOF, CultureInfo.InvariantCulture))
            {
                var row = new List<string>();
                for (int i = 0; i < fieldCount; i++)
                {
                    dynamic field = recordset.Fields[i];
                    row.Add(FormatValue(field.Value));
                }

                rows.Add(row);
                recordset.MoveNext();
            }

            return new AccessQueryResult(title, columns, rows);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("Ошибка выполнения Access-запроса: " + ex.Message, ex);
        }
        finally
        {
            CloseRecordset(recordset);
            AccessDatabaseCreator.ReleaseComObject(recordset);
            CloseConnection(connection);
            AccessDatabaseCreator.ReleaseComObject(connection);
        }
    }

    private void CheckDatabaseFile()
    {
        if (!File.Exists(_dbPath))
        {
            throw new FileNotFoundException("Файл базы данных MS Access не найден.", _dbPath);
        }
    }

    private static string FormatValue(object? value)
    {
        if (value == null || value == DBNull.Value) return string.Empty;
        if (value is double d) return d.ToString("0.##", CultureInfo.InvariantCulture);
        if (value is float f) return f.ToString("0.##", CultureInfo.InvariantCulture);
        if (value is decimal m) return m.ToString("0.##", CultureInfo.InvariantCulture);
        return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
    }

    private static void CloseRecordset(dynamic? recordset)
    {
        if (recordset == null) return;

        try
        {
            if (recordset.State != 0) recordset.Close();
        }
        catch (COMException)
        {
        }
    }

    private static void CloseConnection(dynamic? connection)
    {
        if (connection == null) return;

        try
        {
            if (connection.State != 0) connection.Close();
        }
        catch (COMException)
        {
        }
    }
}
