using System.Globalization;
using System.Runtime.InteropServices;

namespace ElectricityProductionApp;

public static class AccessDatabaseCreator
{
    private static readonly string[] Providers =
    {
        "Microsoft.ACE.OLEDB.16.0",
        "Microsoft.ACE.OLEDB.12.0"
    };

    public static void Create(string dbPath, bool overwrite)
    {
        string fullPath = Path.GetFullPath(dbPath);
        string? directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);

        if (File.Exists(fullPath))
        {
            if (!overwrite) return;
            File.Delete(fullPath);
        }

        Exception? lastError = null;
        foreach (string provider in Providers)
        {
            try
            {
                CreateWithProvider(fullPath, provider);
                return;
            }
            catch (Exception ex) when (ex is COMException or InvalidOperationException)
            {
                lastError = ex;
                if (File.Exists(fullPath)) File.Delete(fullPath);
            }
        }

        throw new InvalidOperationException(
            "Не удалось создать файл MS Access. Установите Microsoft Access или Microsoft Access Database Engine.",
            lastError);
    }

    private static void CreateWithProvider(string dbPath, string provider)
    {
        string createConnectionString = $"Provider={provider};Data Source={dbPath};";
        string connectionString = BuildConnectionString(provider, dbPath);
        dynamic? catalog = null;
        dynamic? connection = null;

        try
        {
            catalog = CreateComObject("ADOX.Catalog");
            dynamic? createdConnection = catalog.Create(createConnectionString);
            CloseConnection(createdConnection);
            ReleaseComObject(createdConnection);
            ReleaseComObject(catalog);
            catalog = null;

            connection = CreateComObject("ADODB.Connection");
            connection.Open(connectionString);

            ExecuteNonQuery(connection, "CREATE TABLE Regions (RegionId AUTOINCREMENT CONSTRAINT PK_Regions PRIMARY KEY, RegionName TEXT(100) NOT NULL)");
            ExecuteNonQuery(connection, "CREATE UNIQUE INDEX UX_Regions_RegionName ON Regions (RegionName)");
            ExecuteNonQuery(connection, "CREATE TABLE EnergyProduction (ProductionId AUTOINCREMENT CONSTRAINT PK_EnergyProduction PRIMARY KEY, RegionId LONG NOT NULL, CountryName TEXT(100) NOT NULL, Production2010 DOUBLE NOT NULL, Production2015 DOUBLE NOT NULL)");
            ExecuteNonQuery(connection, "ALTER TABLE EnergyProduction ADD CONSTRAINT FK_EnergyProduction_Regions FOREIGN KEY (RegionId) REFERENCES Regions(RegionId)");
            ExecuteNonQuery(connection, "CREATE UNIQUE INDEX UX_EnergyProduction_CountryName ON EnergyProduction (CountryName)");
            ExecuteNonQuery(connection, "CREATE INDEX IX_EnergyProduction_RegionId ON EnergyProduction (RegionId)");

            foreach (string region in CsvRepository.CreateDefaultRecords().Select(record => record.Region).Distinct())
            {
                ExecuteNonQuery(connection, $"INSERT INTO Regions (RegionName) VALUES ('{Escape(region)}')");
            }

            foreach (ElectricityRecord record in CsvRepository.CreateDefaultRecords())
            {
                string p2010 = record.Year2010.ToString(CultureInfo.InvariantCulture);
                string p2015 = record.Year2015.ToString(CultureInfo.InvariantCulture);
                ExecuteNonQuery(connection,
                    "INSERT INTO EnergyProduction (RegionId, CountryName, Production2010, Production2015) " +
                    $"SELECT RegionId, '{Escape(record.Country)}', {p2010}, {p2015} FROM Regions WHERE RegionName='{Escape(record.Region)}'");
            }

            CloseConnection(connection);
            ReleaseComObject(connection);
            connection = null;

            TryCreateSavedQueries(dbPath);
        }
        finally
        {
            CloseConnection(connection);
            ReleaseComObject(connection);
            ReleaseComObject(catalog);
        }
    }

    private static void TryCreateSavedQueries(string dbPath)
    {
        try
        {
            CreateSavedQueries(dbPath);
        }
        catch (Exception ex) when (ex is COMException or InvalidOperationException)
        {
        }
    }

    private static void CreateSavedQueries(string dbPath)
    {
        dynamic? access = null;
        dynamic? database = null;

        try
        {
            access = CreateComObject("Access.Application");
            access.OpenCurrentDatabase(dbPath);
            database = access.CurrentDb();

            CreateQuery(database,
                "qryTopCountry2015",
                "SELECT TOP 1 EnergyProduction.CountryName, EnergyProduction.Production2015 FROM EnergyProduction ORDER BY EnergyProduction.Production2015 DESC;");

            CreateQuery(database,
                "qryCountriesOver70In2015",
                "SELECT Regions.RegionName, EnergyProduction.CountryName, EnergyProduction.Production2015 FROM Regions INNER JOIN EnergyProduction ON Regions.RegionId = EnergyProduction.RegionId WHERE EnergyProduction.Production2015 > 70 ORDER BY EnergyProduction.CountryName;");

            CreateQuery(database,
                "qryCountriesNotOver100In2010",
                "SELECT Regions.RegionName, EnergyProduction.CountryName, EnergyProduction.Production2010 FROM Regions INNER JOIN EnergyProduction ON Regions.RegionId = EnergyProduction.RegionId WHERE EnergyProduction.Production2010 <= 100 ORDER BY EnergyProduction.CountryName;");
        }
        finally
        {
            ReleaseComObject(database);
            CloseAccess(access);
            ReleaseComObject(access);
        }
    }

    private static void CreateQuery(dynamic database, string name, string sql)
    {
        dynamic? query = null;

        try
        {
            try
            {
                database.QueryDefs.Delete(name);
            }
            catch (COMException)
            {
            }

            query = database.CreateQueryDef(name, sql);
        }
        finally
        {
            ReleaseComObject(query);
        }
    }

    internal static dynamic OpenConnection(string dbPath)
    {
        Exception? lastError = null;

        foreach (string provider in Providers)
        {
            dynamic? connection = null;
            try
            {
                connection = CreateComObject("ADODB.Connection");
                connection.Open(BuildConnectionString(provider, dbPath));
                return connection;
            }
            catch (Exception ex) when (ex is COMException or InvalidOperationException)
            {
                lastError = ex;
                CloseConnection(connection);
                ReleaseComObject(connection);
            }
        }

        throw new InvalidOperationException(
            "Не удалось открыть базу MS Access. Установите Microsoft Access или Microsoft Access Database Engine.",
            lastError);
    }

    internal static void ReleaseComObject(object? comObject)
    {
        if (comObject != null && Marshal.IsComObject(comObject))
        {
            Marshal.FinalReleaseComObject(comObject);
        }
    }

    private static dynamic CreateComObject(string progId)
    {
        Type? type = Type.GetTypeFromProgID(progId);
        if (type == null) throw new InvalidOperationException($"COM-объект {progId} недоступен.");
        return Activator.CreateInstance(type) ?? throw new InvalidOperationException($"COM-объект {progId} не создан.");
    }

    private static string BuildConnectionString(string provider, string dbPath)
    {
        return $"Provider={provider};Data Source={dbPath};Persist Security Info=False;";
    }

    private static void ExecuteNonQuery(dynamic connection, string sql)
    {
        connection.Execute(sql);
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

    private static void CloseAccess(dynamic? access)
    {
        if (access == null) return;

        try
        {
            access.CloseCurrentDatabase();
            access.Quit();
        }
        catch (COMException)
        {
        }
    }

    private static string Escape(string value)
    {
        return value.Replace("'", "''");
    }
}
