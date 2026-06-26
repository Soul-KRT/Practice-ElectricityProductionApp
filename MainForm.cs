using System.ComponentModel;
using System.Drawing;
using System.Drawing.Printing;
using System.Globalization;
using System.Text;
using System.Windows.Forms;

namespace ElectricityProductionApp;

public class MainForm : Form
{
    private readonly string _projectDirectory;
    private readonly string _dataDirectory;
    private readonly string _csvPath;
    private readonly string _accessPath;
    private readonly string _sqlPath;
    private readonly CsvRepository _repository;
    private readonly AccessRepository _accessRepository;

    private BindingList<ElectricityRecord> _records = new();
    private string _lastResultText = string.Empty;
    private string _resultFileText = string.Empty;
    private bool _showChart;

    private DataGridView _grid = new();
    private TextBox _txtRegion = new();
    private TextBox _txtCountry = new();
    private NumericUpDown _nud2010 = new();
    private NumericUpDown _nud2015 = new();
    private ComboBox _cmbSort = new();
    private TextBox _txtResult = new();
    private Panel _chartPanel = new();

    public MainForm()
    {
        _projectDirectory = GetProjectDirectory();
        _dataDirectory = Path.Combine(_projectDirectory, "data");
        _csvPath = Path.Combine(_dataDirectory, "electricity.csv");
        _accessPath = Path.Combine(_dataDirectory, "EnergyProduction.accdb");
        _sqlPath = Path.Combine(_projectDirectory, "access_queries.sql");
        _repository = new CsvRepository(_csvPath);
        _accessRepository = new AccessRepository(_accessPath);

        Text = "Вариант 13 - производство электроэнергии";
        Width = 1230;
        Height = 790;
        MinimumSize = new Size(1080, 650);
        StartPosition = FormStartPosition.CenterScreen;

        BuildInterface();
        LoadRecords();
        InitializeAccessDatabase();
    }

    private void BuildInterface()
    {
        var menu = new MenuStrip();

        var fileMenu = new ToolStripMenuItem("Файл");
        fileMenu.DropDownItems.Add("Создать CSV-файл", null, (_, _) => CreateFile());
        fileMenu.DropDownItems.Add("Печать исходного CSV-файла", null, (_, _) => PrintSourceFile());
        fileMenu.DropDownItems.Add("Сохранить результирующий файл", null, (_, _) => SaveResultToFile());
        fileMenu.DropDownItems.Add("Печать результирующего файла", null, (_, _) => PrintResultFile());
        fileMenu.DropDownItems.Add("Выход", null, (_, _) => Close());

        var recordsMenu = new ToolStripMenuItem("CSV-записи");
        recordsMenu.DropDownItems.Add("Добавить запись", null, (_, _) => AddRecord());
        recordsMenu.DropDownItems.Add("Корректировать запись", null, (_, _) => UpdateRecord());
        recordsMenu.DropDownItems.Add("Удалить выбранные записи", null, (_, _) => DeleteSelectedRecords());
        recordsMenu.DropDownItems.Add("Упорядочить записи", null, (_, _) => SortRecords());

        var csvTaskMenu = new ToolStripMenuItem("CSV-расчеты");
        csvTaskMenu.DropDownItems.Add("Итоговые показатели", null, (_, _) => CalculateTotals());
        csvTaskMenu.DropDownItems.Add("Задание 1: максимум за 2015 г.", null, (_, _) => ExecuteCsvTask1());
        csvTaskMenu.DropDownItems.Add("Задание 2: 2015 г. > 70", null, (_, _) => ExecuteCsvTask2());
        csvTaskMenu.DropDownItems.Add("Задание 3: 2010 г. <= 100", null, (_, _) => ExecuteCsvTask3());
        csvTaskMenu.DropDownItems.Add("График сводных данных", null, (_, _) => ShowSummaryChart());

        var accessMenu = new ToolStripMenuItem("MS Access");
        accessMenu.DropDownItems.Add("Создать файл БД", null, (_, _) => RecreateAccessDatabase());
        accessMenu.DropDownItems.Add("Показать таблицу", null, (_, _) => ShowAccessTable());
        accessMenu.DropDownItems.Add("Запрос 1: максимум 2015", null, (_, _) => ExecuteAccessTask1());
        accessMenu.DropDownItems.Add("Запрос 2: 2015 г. > 70", null, (_, _) => ExecuteAccessTask2());
        accessMenu.DropDownItems.Add("Запрос 3: 2010 г. <= 100", null, (_, _) => ExecuteAccessTask3());
        accessMenu.DropDownItems.Add("Показать SQL Access", null, (_, _) => ShowAccessSql());

        menu.Items.AddRange(new ToolStripItem[] { fileMenu, recordsMenu, csvTaskMenu, accessMenu });
        MainMenuStrip = menu;
        Controls.Add(menu);

        _grid = new DataGridView
        {
            Left = 12,
            Top = 38,
            Width = 735,
            Height = 330,
            AutoGenerateColumns = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = true,
            AllowUserToAddRows = false,
            ReadOnly = true,
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Регион", DataPropertyName = "Region", Width = 145 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Страна", DataPropertyName = "Country", Width = 120 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "2010", DataPropertyName = "Year2010", Width = 75 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "2015", DataPropertyName = "Year2015", Width = 75 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Прирост", DataPropertyName = "Increase", Width = 80 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Прирост, %", DataPropertyName = "IncreasePercent", Width = 95 });
        _grid.SelectionChanged += (_, _) => FillInputsFromSelectedRecord();
        Controls.Add(_grid);

        var inputBox = new GroupBox
        {
            Text = "Часть 2. CSV-файл: запись",
            Left = 765,
            Top = 38,
            Width = 425,
            Height = 190,
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };
        Controls.Add(inputBox);

        AddLabel(inputBox, "Регион", 14, 32);
        _txtRegion = AddTextBox(inputBox, 110, 28, 280);
        AddLabel(inputBox, "Страна", 14, 66);
        _txtCountry = AddTextBox(inputBox, 110, 62, 280);
        AddLabel(inputBox, "2010", 14, 100);
        _nud2010 = AddNumber(inputBox, 110, 96);
        AddLabel(inputBox, "2015", 14, 134);
        _nud2015 = AddNumber(inputBox, 110, 130);

        AddButton(inputBox, "Добавить", 14, 156, AddRecord, 105, 26);
        AddButton(inputBox, "Изменить", 132, 156, UpdateRecord, 105, 26);
        AddButton(inputBox, "Удалить", 250, 156, DeleteSelectedRecords, 105, 26);

        var csvBox = new GroupBox
        {
            Text = "Часть 2. CSV-файл: режимы",
            Left = 765,
            Top = 238,
            Width = 425,
            Height = 145,
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };
        Controls.Add(csvBox);

        _cmbSort = new ComboBox
        {
            Left = 14,
            Top = 28,
            Width = 145,
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        _cmbSort.Items.AddRange(new object[] { "Страна", "Регион", "2010", "2015" });
        _cmbSort.SelectedIndex = 0;
        csvBox.Controls.Add(_cmbSort);
        AddButton(csvBox, "Сортировать", 170, 26, SortRecords, 110);
        AddButton(csvBox, "Задание 1", 14, 68, ExecuteCsvTask1, 90);
        AddButton(csvBox, "Задание 2", 112, 68, ExecuteCsvTask2, 90);
        AddButton(csvBox, "Задание 3", 210, 68, ExecuteCsvTask3, 90);
        AddButton(csvBox, "Итоги", 308, 68, CalculateTotals, 80);
        AddButton(csvBox, "График", 14, 105, ShowSummaryChart, 90);
        AddButton(csvBox, "Сохранить", 112, 105, SaveResultToFile, 90);
        AddButton(csvBox, "Печать", 210, 105, PrintResultFile, 90);

        var accessBox = new GroupBox
        {
            Text = "Часть 1. MS Access: БД и запросы",
            Left = 765,
            Top = 393,
            Width = 425,
            Height = 155,
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };
        Controls.Add(accessBox);

        AddButton(accessBox, "Создать БД", 14, 28, RecreateAccessDatabase, 110);
        AddButton(accessBox, "Таблица", 132, 28, ShowAccessTable, 90);
        AddButton(accessBox, "SQL", 230, 28, ShowAccessSql, 70);
        AddButton(accessBox, "Запрос 1", 14, 70, ExecuteAccessTask1, 90);
        AddButton(accessBox, "Запрос 2", 112, 70, ExecuteAccessTask2, 90);
        AddButton(accessBox, "Запрос 3", 210, 70, ExecuteAccessTask3, 90);
        AddButton(accessBox, "Путь к БД", 308, 70, ShowAccessPath, 90);

        _txtResult = new TextBox
        {
            Left = 12,
            Top = 382,
            Width = 735,
            Height = 345,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            ReadOnly = true,
            Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom | AnchorStyles.Top
        };
        Controls.Add(_txtResult);

        _chartPanel = new Panel
        {
            Left = 765,
            Top = 560,
            Width = 425,
            Height = 167,
            BorderStyle = BorderStyle.FixedSingle,
            Anchor = AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom
        };
        _chartPanel.Paint += DrawChart;
        Controls.Add(_chartPanel);
    }

    private void LoadRecords()
    {
        _records = new BindingList<ElectricityRecord>(_repository.Load());
        _grid.DataSource = _records;
        WriteResult("CSV-файл загружен: " + _csvPath);
    }

    private void InitializeAccessDatabase()
    {
        if (_accessRepository.Exists())
        {
            WriteResult("CSV-файл загружен: " + _csvPath + Environment.NewLine +
                        "MS Access БД готова: " + _accessPath);
            return;
        }

        WriteResult("CSV-файл загружен: " + _csvPath + Environment.NewLine +
                    "MS Access БД не найдена: " + _accessPath + Environment.NewLine +
                    "Для создания файла нажмите MS Access -> Создать файл БД.");
    }

    private void CreateFile()
    {
        var defaults = CsvRepository.CreateDefaultRecords();
        _repository.Save(defaults);
        _records = new BindingList<ElectricityRecord>(defaults);
        _grid.DataSource = _records;
        WriteResult("Создан CSV-файл с исходными данными по варианту 13." + Environment.NewLine + _csvPath);
    }

    private void PrintSourceFile()
    {
        PrintText("Исходный CSV-файл", BuildRecordsText(_records));
    }

    private void AddRecord()
    {
        if (!TryReadInputRecord(out ElectricityRecord record)) return;
        if (_records.Any(item => item.Country.Equals(record.Country, StringComparison.OrdinalIgnoreCase)))
        {
            MessageBox.Show("Страна с таким названием уже есть в CSV-файле.");
            return;
        }

        _records.Add(record);
        SaveCurrentRecords();
        WriteResult("Запись добавлена и сохранена в CSV-файл.");
    }

    private void UpdateRecord()
    {
        if (_grid.SelectedRows.Count == 0)
        {
            MessageBox.Show("Выберите запись для корректировки.");
            return;
        }

        if (!TryReadInputRecord(out ElectricityRecord record)) return;
        int index = _grid.SelectedRows[0].Index;
        if (index < 0 || index >= _records.Count) return;

        _records[index] = record;
        SaveCurrentRecords();
        WriteResult("Выбранная запись откорректирована.");
    }

    private void DeleteSelectedRecords()
    {
        if (_grid.SelectedRows.Count == 0)
        {
            MessageBox.Show("Выберите одну или несколько записей для удаления.");
            return;
        }

        var indexes = _grid.SelectedRows.Cast<DataGridViewRow>()
            .Select(row => row.Index)
            .Where(index => index >= 0 && index < _records.Count)
            .OrderByDescending(index => index)
            .ToList();

        foreach (int index in indexes)
        {
            _records.RemoveAt(index);
        }

        SaveCurrentRecords();
        WriteResult($"Удалено записей из CSV-файла: {indexes.Count}.");
    }

    private void SortRecords()
    {
        IEnumerable<ElectricityRecord> sorted = _cmbSort.Text switch
        {
            "Регион" => _records.OrderBy(record => record.Region).ThenBy(record => record.Country),
            "2010" => _records.OrderByDescending(record => record.Year2010),
            "2015" => _records.OrderByDescending(record => record.Year2015),
            _ => _records.OrderBy(record => record.Country)
        };

        _records = new BindingList<ElectricityRecord>(sorted.Select(record => record.Clone()).ToList());
        _grid.DataSource = _records;
        SaveCurrentRecords();
        WriteResult("CSV-записи упорядочены по полю: " + _cmbSort.Text + ".");
    }

    private void CalculateTotals()
    {
        _resultFileText = BuildResultFileText();
        SaveResultText(_resultFileText);
        WriteResult(_resultFileText + Environment.NewLine + Environment.NewLine +
                    "Результирующий файл сформирован: " + GetResultPath());
        _showChart = true;
        _chartPanel.Invalidate();
    }

    private void ExecuteCsvTask1()
    {
        ElectricityRecord? max = QueryService.CountryWithMax2015(_records);
        if (max == null)
        {
            WriteResult("Нет данных для выполнения задания 1.");
            return;
        }

        WriteResult("CSV. Задание 1. Страна, которая в 2015 г. произвела больше всех электроэнергии:" +
                    Environment.NewLine + $"{max.Country} - {max.Year2015:0.##} млрд кВт/час.");
    }

    private void ExecuteCsvTask2()
    {
        var result = QueryService.CountriesWith2015MoreThan70(_records);
        WriteResult("CSV. Задание 2. Страны, в которых в 2015 г. производство электроэнергии превысило 70 млрд кВт/час:" +
                    Environment.NewLine + BuildRecordsText(result));
    }

    private void ExecuteCsvTask3()
    {
        var result = QueryService.CountriesWith2010NotMoreThan100(_records);
        WriteResult("CSV. Задание 3. В исходном варианте указан 2000 г., но в таблице есть годы 2010 и 2015. " +
                    "Поэтому запрос выполнен по 2010 г.: страны, где производство не превышало 100 млрд кВт/час." +
                    Environment.NewLine + BuildRecordsText(result));
    }

    private void ShowSummaryChart()
    {
        _showChart = true;
        _chartPanel.Invalidate();
        ShowChartWindow();
        WriteResult("Построена сводная диаграмма: сравнение производства электроэнергии за 2010 и 2015 годы по странам.");
    }

    private void ShowChartWindow()
    {
        var form = new Form
        {
            Text = "Сводная диаграмма производства электроэнергии",
            Width = 900,
            Height = 520,
            StartPosition = FormStartPosition.CenterParent
        };

        var panel = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.White
        };
        panel.Paint += DrawChart;
        form.Controls.Add(panel);
        form.Show(this);
    }

    private void RecreateAccessDatabase()
    {
        RunAccessAction(() =>
        {
            _accessRepository.Recreate();
            return "Файл MS Access создан заново:" + Environment.NewLine + _accessPath + Environment.NewLine +
                   Environment.NewLine + _accessRepository.LoadAll().ToDisplayText();
        });
    }

    private void ShowAccessTable()
    {
        RunAccessQuery(() => _accessRepository.LoadAll());
    }

    private void ExecuteAccessTask1()
    {
        RunAccessQuery(() => _accessRepository.QueryTopCountry2015());
    }

    private void ExecuteAccessTask2()
    {
        RunAccessQuery(() => _accessRepository.QueryCountriesOver70In2015());
    }

    private void ExecuteAccessTask3()
    {
        RunAccessQuery(() => _accessRepository.QueryCountriesNotOver100In2010());
    }

    private void ShowAccessSql()
    {
        if (!File.Exists(_sqlPath))
        {
            WriteResult("Файл access_queries.sql не найден.");
            return;
        }

        WriteResult(File.ReadAllText(_sqlPath, Encoding.UTF8));
    }

    private void ShowAccessPath()
    {
        WriteResult("Файл MS Access:" + Environment.NewLine + _accessPath + Environment.NewLine +
                    "Файл SQL-запросов:" + Environment.NewLine +
                    _sqlPath);
    }

    private void SaveResultToFile()
    {
        if (string.IsNullOrWhiteSpace(_resultFileText))
        {
            MessageBox.Show("Сначала выполните расчет итоговых показателей.");
            return;
        }

        SaveResultText(_resultFileText);
        MessageBox.Show("Результирующий файл сохранен: " + GetResultPath());
    }

    private void PrintResultFile()
    {
        if (string.IsNullOrWhiteSpace(_resultFileText))
        {
            MessageBox.Show("Сначала выполните расчет итоговых показателей.");
            return;
        }

        PrintText("Результирующий файл", _resultFileText);
    }

    private string GetResultPath()
    {
        return Path.Combine(_dataDirectory, "result.txt");
    }

    private void SaveResultText(string text)
    {
        Directory.CreateDirectory(_dataDirectory);
        File.WriteAllText(GetResultPath(), text, Encoding.UTF8);
    }

    private void SaveCurrentRecords()
    {
        _repository.Save(_records);
        _grid.Refresh();
        _chartPanel.Invalidate();
    }

    private bool TryReadInputRecord(out ElectricityRecord record)
    {
        record = new ElectricityRecord();
        if (string.IsNullOrWhiteSpace(_txtRegion.Text) || string.IsNullOrWhiteSpace(_txtCountry.Text))
        {
            MessageBox.Show("Заполните регион и страну.");
            return false;
        }

        record.Region = _txtRegion.Text.Trim();
        record.Country = _txtCountry.Text.Trim();
        record.Year2010 = _nud2010.Value;
        record.Year2015 = _nud2015.Value;
        return true;
    }

    private void FillInputsFromSelectedRecord()
    {
        if (_grid.SelectedRows.Count == 0) return;
        if (_grid.SelectedRows[0].DataBoundItem is not ElectricityRecord record) return;

        _txtRegion.Text = record.Region;
        _txtCountry.Text = record.Country;
        _nud2010.Value = ClampToNumeric(record.Year2010, _nud2010);
        _nud2015.Value = ClampToNumeric(record.Year2015, _nud2015);
    }

    private static decimal ClampToNumeric(decimal value, NumericUpDown numeric)
    {
        if (value < numeric.Minimum) return numeric.Minimum;
        if (value > numeric.Maximum) return numeric.Maximum;
        return value;
    }

    private static string BuildRecordsText(IEnumerable<ElectricityRecord> records)
    {
        var sb = new StringBuilder();
        sb.AppendLine("Регион; Страна; 2010; 2015; Прирост; Прирост, %");
        foreach (ElectricityRecord record in records)
        {
            sb.AppendLine($"{record.Region}; {record.Country}; {record.Year2010:0.##}; {record.Year2015:0.##}; {record.Increase:0.##}; {record.IncreasePercent:0.##}%");
        }

        return sb.ToString();
    }

    private string BuildResultFileText()
    {
        return "Результирующий файл по варианту 13" + Environment.NewLine +
               "Производство электроэнергии в развитых странах мира" + Environment.NewLine +
               Environment.NewLine +
               BuildRecordsText(_records) + Environment.NewLine +
               QueryService.BuildSummary(_records);
    }

    private void RunAccessQuery(Func<AccessQueryResult> queryFactory)
    {
        RunAccessAction(() => queryFactory().ToDisplayText());
    }

    private void RunAccessAction(Func<string> action)
    {
        try
        {
            WriteResult(action());
        }
        catch (Exception ex)
        {
            string message = "MS Access операция не выполнена: " + ex.Message;
            WriteResult(message);
            MessageBox.Show(message);
        }
    }

    private void WriteResult(string text)
    {
        _lastResultText = text;
        _txtResult.Text = text;
    }

    private void PrintText(string title, string text)
    {
        var printDocument = new PrintDocument { DocumentName = title };
        printDocument.PrintPage += (_, eventArgs) =>
        {
            if (eventArgs.Graphics == null) return;

            using var font = new Font("Times New Roman", 12);
            eventArgs.Graphics.DrawString(title + Environment.NewLine + Environment.NewLine + text,
                font,
                Brushes.Black,
                eventArgs.MarginBounds);
        };

        using var preview = new PrintPreviewDialog
        {
            Document = printDocument,
            Width = 900,
            Height = 700
        };
        preview.ShowDialog(this);
    }

    private void DrawChart(object? sender, PaintEventArgs e)
    {
        Panel panel = sender as Panel ?? _chartPanel;
        Graphics g = e.Graphics;
        g.Clear(Color.White);
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

        using var titleFont = new Font("Times New Roman", 11, FontStyle.Bold);
        using var textFont = new Font("Times New Roman", 8);
        using var smallFont = new Font("Times New Roman", 7);
        using var blueBrush = new SolidBrush(Color.FromArgb(155, 185, 225));
        using var orangeBrush = new SolidBrush(Color.FromArgb(230, 180, 150));
        using var borderPen = new Pen(Color.FromArgb(70, 70, 70));

        string title = "Сводная диаграмма производства электроэнергии";
        SizeF titleSize = g.MeasureString(title, titleFont);
        g.DrawString(title, titleFont, Brushes.Black, (panel.Width - titleSize.Width) / 2, 8);

        if (!_showChart || _records.Count == 0)
        {
            g.DrawString("Нажмите кнопку 'График'.", textFont, Brushes.Gray, 10, 40);
            return;
        }

        int legendLeft = Math.Max(10, panel.Width - 115);
        g.FillRectangle(blueBrush, legendLeft, 36, 24, 14);
        g.DrawRectangle(borderPen, legendLeft, 36, 24, 14);
        g.DrawString("2010", textFont, Brushes.Black, legendLeft + 32, 34);
        g.FillRectangle(orangeBrush, legendLeft, 58, 24, 14);
        g.DrawRectangle(borderPen, legendLeft, 58, 24, 14);
        g.DrawString("2015", textFont, Brushes.Black, legendLeft + 32, 56);

        int plotLeft = 35;
        int plotTop = 72;
        int plotWidth = Math.Max(80, panel.Width - 55);
        int plotHeight = Math.Max(60, panel.Height - 108);
        decimal max = _records.Max(record => Math.Max(record.Year2010, record.Year2015));
        if (max <= 0) max = 1;

        using var axisPen = new Pen(Color.Black, 2);
        g.DrawLine(axisPen, plotLeft, plotTop, plotLeft, plotTop + plotHeight);
        g.DrawLine(axisPen, plotLeft, plotTop + plotHeight, plotLeft + plotWidth, plotTop + plotHeight);

        int groupWidth = plotWidth / Math.Max(1, _records.Count);
        int barWidth = Math.Max(8, Math.Min(22, (groupWidth - 12) / 2));
        for (int i = 0; i < _records.Count; i++)
        {
            ElectricityRecord record = _records[i];
            int groupLeft = plotLeft + 6 + i * groupWidth;
            int x2010 = groupLeft;
            int x2015 = groupLeft + barWidth + 4;

            DrawBar(g, record.Year2010, max, x2010, plotTop, plotHeight, barWidth, blueBrush, borderPen, smallFont);
            DrawBar(g, record.Year2015, max, x2015, plotTop, plotHeight, barWidth, orangeBrush, borderPen, smallFont);

            SizeF countrySize = g.MeasureString(record.Country, textFont);
            float countryX = groupLeft + barWidth - countrySize.Width / 2;
            g.DrawString(record.Country, textFont, Brushes.Black, countryX, plotTop + plotHeight + 4);
        }
    }

    private static void DrawBar(Graphics g, decimal value, decimal max, int x, int plotTop, int plotHeight, int width, Brush brush, Pen borderPen, Font font)
    {
        int height = (int)(plotHeight * value / max);
        int y = plotTop + plotHeight - height;
        g.FillRectangle(brush, x, y, width, height);
        g.DrawRectangle(borderPen, x, y, width, height);

        string text = value.ToString("0.#", CultureInfo.InvariantCulture);
        SizeF textSize = g.MeasureString(text, font);
        g.DrawString(text, font, Brushes.Black, x + width / 2f - textSize.Width / 2f, y - textSize.Height - 1);
    }

    private static void AddLabel(Control parent, string text, int left, int top)
    {
        parent.Controls.Add(new Label { Text = text, Left = left, Top = top + 4, Width = 90 });
    }

    private static TextBox AddTextBox(Control parent, int left, int top, int width)
    {
        var textBox = new TextBox { Left = left, Top = top, Width = width };
        parent.Controls.Add(textBox);
        return textBox;
    }

    private static NumericUpDown AddNumber(Control parent, int left, int top)
    {
        var number = new NumericUpDown
        {
            Left = left,
            Top = top,
            Width = 120,
            DecimalPlaces = 1,
            Minimum = 0,
            Maximum = 10000,
            Increment = 0.1m
        };
        parent.Controls.Add(number);
        return number;
    }

    private static Button AddButton(Control parent, string text, int left, int top, Action action, int width = 105, int height = 30)
    {
        var button = new Button
        {
            Text = text,
            Left = left,
            Top = top,
            Width = width,
            Height = height
        };
        button.Click += (_, _) => action();
        parent.Controls.Add(button);
        return button;
    }

    private static string GetProjectDirectory()
    {
        DirectoryInfo? directory = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);

        while (directory != null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "ElectricityProductionApp.csproj")))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        return AppDomain.CurrentDomain.BaseDirectory;
    }
}
