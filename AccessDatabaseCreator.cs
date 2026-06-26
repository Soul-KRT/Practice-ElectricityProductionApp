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

            TryCreateAccessObjects(dbPath);
        }
        finally
        {
            CloseConnection(connection);
            ReleaseComObject(connection);
            ReleaseComObject(catalog);
        }
    }

    private static void TryCreateAccessObjects(string dbPath)
    {
        try
        {
            CreateAccessObjects(dbPath);
        }
        catch (Exception ex) when (ex is COMException or InvalidOperationException)
        {
        }
    }

    private static void CreateAccessObjects(string dbPath)
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

            CreateForms(access);
            SetDatabaseProperty(database, "StartupForm", "frmMain");
        }
        finally
        {
            ReleaseComObject(database);
            CloseAccess(access);
            ReleaseComObject(access);
        }
    }

    private static void CreateForms(dynamic access)
    {
        DeleteForm(access, "frmMain");
        DeleteForm(access, "frmData");
        DeleteForm(access, "frmAnalytics");

        CreateMainForm(access);
        CreateDataForm(access);
        CreateAnalyticsForm(access);
    }

    private static void CreateMainForm(dynamic access)
    {
        dynamic form = access.CreateForm();
        string formName = form.Name;
        form.Caption = "Главная форма базы данных";
        form.Width = 7600;
        form.Section(0).Height = 4300;

        AddLabel(access, formName, "Производство электроэнергии", 900, 400, 5600, 450, 18, true);
        AddLabel(access, formName, "Вариант 13", 2900, 900, 1800, 300, 12, false);

        dynamic dataButton = AddButton(access, formName, "cmdOpenData", "Ввод и редактирование данных", 1900, 1600, 3600, 450);
        AddButtonEvent(form, dataButton, "    DoCmd.OpenForm \"frmData\"");

        dynamic analyticsButton = AddButton(access, formName, "cmdOpenAnalytics", "Аналитические запросы", 1900, 2200, 3600, 450);
        AddButtonEvent(form, analyticsButton, "    DoCmd.OpenForm \"frmAnalytics\"");

        dynamic exitButton = AddButton(access, formName, "cmdExit", "Выход", 1900, 2800, 3600, 450);
        AddButtonEvent(form, exitButton, "    DoCmd.Quit");

        SaveForm(access, formName, "frmMain");
    }

    private static void CreateDataForm(dynamic access)
    {
        dynamic form = access.CreateForm();
        string formName = form.Name;
        form.Caption = "Ввод и редактирование данных";
        form.RecordSource = "EnergyProduction";
        form.Width = 8500;
        form.Section(0).Height = 5200;
        form.NavigationButtons = true;
        form.AllowAdditions = true;
        form.AllowEdits = true;
        form.AllowDeletions = true;

        AddLabel(access, formName, "Данные о производстве электроэнергии", 850, 300, 6500, 400, 16, true);

        AddLabel(access, formName, "Страна", 850, 1100, 1800, 300, 10, false);
        AddTextBox(access, formName, "txtCountry", "CountryName", 2900, 1050, 3000, 350);

        AddLabel(access, formName, "Регион", 850, 1600, 1800, 300, 10, false);
        dynamic combo = access.CreateControl(formName, 111, 0, null, null, 2900, 1550, 3000, 350);
        combo.Name = "cmbRegion";
        combo.ControlSource = "RegionId";
        combo.RowSourceType = "Table/Query";
        combo.RowSource = "SELECT RegionId, RegionName FROM Regions ORDER BY RegionName";
        combo.BoundColumn = 1;
        combo.ColumnCount = 2;
        combo.ColumnWidths = "0;3000";
        combo.LimitToList = true;

        AddLabel(access, formName, "Показатель за 2010 г.", 850, 2100, 1800, 300, 10, false);
        AddTextBox(access, formName, "txt2010", "Production2010", 2900, 2050, 1600, 350);

        AddLabel(access, formName, "Показатель за 2015 г.", 850, 2600, 1800, 300, 10, false);
        AddTextBox(access, formName, "txt2015", "Production2015", 2900, 2550, 1600, 350);

        dynamic newButton = AddButton(access, formName, "cmdNew", "Новая запись", 850, 3400, 1700, 420);
        AddButtonEvent(form, newButton, "    DoCmd.GoToRecord , , acNewRec");

        dynamic saveButton = AddButton(access, formName, "cmdSave", "Сохранить", 2750, 3400, 1700, 420);
        AddButtonEvent(form, saveButton, "    If Me.Dirty Then" + Environment.NewLine + "        DoCmd.RunCommand acCmdSaveRecord" + Environment.NewLine + "    End If");

        dynamic deleteButton = AddButton(access, formName, "cmdDelete", "Удалить", 4650, 3400, 1700, 420);
        AddButtonEvent(form, deleteButton, "    If Not Me.NewRecord Then" + Environment.NewLine + "        DoCmd.RunCommand acCmdDeleteRecord" + Environment.NewLine + "    End If");

        dynamic closeButton = AddButton(access, formName, "cmdClose", "Закрыть", 2750, 4000, 1700, 420);
        AddButtonEvent(form, closeButton, "    DoCmd.Close acForm, Me.Name");

        SaveForm(access, formName, "frmData");
    }

    private static void CreateAnalyticsForm(dynamic access)
    {
        dynamic form = access.CreateForm();
        string formName = form.Name;
        form.Caption = "Аналитические запросы";
        form.Width = 8500;
        form.Section(0).Height = 5000;

        AddLabel(access, formName, "Запросы по варианту 13", 1300, 350, 5600, 400, 16, true);
        AddLabel(access, formName, "Каждая кнопка открывает результат запроса в режиме таблицы.", 1000, 850, 6500, 350, 10, false);

        dynamic query1Button = AddButton(access, formName, "cmdQuery1", "1. Максимум за 2015 г.", 1800, 1550, 4500, 450);
        AddButtonEvent(form, query1Button, "    DoCmd.OpenQuery \"qryTopCountry2015\", acViewNormal, acReadOnly");

        dynamic query2Button = AddButton(access, formName, "cmdQuery2", "2. Производство 2015 г. > 70", 1800, 2200, 4500, 450);
        AddButtonEvent(form, query2Button, "    DoCmd.OpenQuery \"qryCountriesOver70In2015\", acViewNormal, acReadOnly");

        dynamic query3Button = AddButton(access, formName, "cmdQuery3", "3. Производство 2010 г. <= 100", 1800, 2850, 4500, 450);
        AddButtonEvent(form, query3Button, "    DoCmd.OpenQuery \"qryCountriesNotOver100In2010\", acViewNormal, acReadOnly");

        dynamic closeButton = AddButton(access, formName, "cmdClose", "Закрыть", 3000, 3650, 2000, 420);
        AddButtonEvent(form, closeButton, "    DoCmd.Close acForm, Me.Name");

        SaveForm(access, formName, "frmAnalytics");
    }

    private static dynamic AddLabel(dynamic access, string formName, string caption, int left, int top, int width, int height, int fontSize, bool bold)
    {
        dynamic label = access.CreateControl(formName, 100, 0, null, null, left, top, width, height);
        label.Caption = caption;
        label.FontSize = fontSize;
        label.FontBold = bold;
        label.TextAlign = 2;
        return label;
    }

    private static dynamic AddTextBox(dynamic access, string formName, string name, string source, int left, int top, int width, int height)
    {
        dynamic textBox = access.CreateControl(formName, 109, 0, null, null, left, top, width, height);
        textBox.Name = name;
        textBox.ControlSource = source;
        return textBox;
    }

    private static dynamic AddButton(dynamic access, string formName, string name, string caption, int left, int top, int width, int height)
    {
        dynamic button = access.CreateControl(formName, 104, 0, null, null, left, top, width, height);
        button.Name = name;
        button.Caption = caption;
        return button;
    }

    private static void AddButtonEvent(dynamic form, dynamic button, string code)
    {
        form.HasModule = true;
        button.OnClick = "[Event Procedure]";
        int line = form.Module.CreateEventProc("Click", button.Name);
        form.Module.InsertLines(line + 1, code);
    }

    private static void SaveForm(dynamic access, string currentName, string finalName)
    {
        access.DoCmd.Save(2, currentName);
        access.DoCmd.Close(2, currentName, 1);
        access.DoCmd.Rename(finalName, 2, currentName);
    }

    private static void DeleteForm(dynamic access, string formName)
    {
        try
        {
            access.DoCmd.DeleteObject(2, formName);
        }
        catch (COMException)
        {
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

    private static void SetDatabaseProperty(dynamic database, string name, string value)
    {
        dynamic? property = null;

        try
        {
            try
            {
                database.Properties.Delete(name);
            }
            catch (COMException)
            {
            }

            property = database.CreateProperty(name, 10, value);
            database.Properties.Append(property);
        }
        finally
        {
            ReleaseComObject(property);
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
