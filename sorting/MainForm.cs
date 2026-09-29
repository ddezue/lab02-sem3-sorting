using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace SortingLab
{
  // Панель с двойной буферизацией — убирает мигание
  public class DoubleBufferedPanel : Panel
  {
    public DoubleBufferedPanel()
    {
      this.DoubleBuffered = true;
      this.SetStyle(ControlStyles.AllPaintingInWmPaint |
                    ControlStyles.UserPaint |
                    ControlStyles.OptimizedDoubleBuffer, true);
      this.UpdateStyles();
    }
  }

  public class MainForm : Form
  {
    // ===== ЭЛЕМЕНТЫ ИНТЕРФЕЙСА =====
    private DataGridView dataInputGrid;
    private DataGridView resultsGrid;
    private DoubleBufferedPanel visualizationPanel;
    private Panel contentPanel;
    private Panel settingsPanel;

    // ===== ЧЕКБОКСЫ =====
    private CheckBox bubbleCheckBox;
    private CheckBox insertionCheckBox;
    private CheckBox shakerCheckBox;
    private CheckBox quickCheckBox;
    private CheckBox bogoCheckBox;
    private CheckBox ascendingCheckBox;
    private CheckBox fractionalCheckBox;
    private CheckBox detailCheckBox;

    // ===== ПОЛЯ ВВОДА =====
    private TextBox countTextBox;
    private TextBox minValueTextBox;
    private TextBox maxValueTextBox;
    private TextBox manualInputTextBox;
    private TextBox maxIterationsTextBox;

    // ===== КНОПКИ / СТАТУС =====
    private Button applyManualButton;
    private Label statusLabel;

    // ===== ДАННЫЕ =====
    private List<double> currentData = new List<double>();
    private const int HistogramLimit = 5000;
    private bool detailedVisualization = true;
    private const int DetailedThreshold = 50;
    private const int MaxRenderedBars = 500;
    private const int MaxRecordedSteps = 5000;

    // ===== ДЛЯ ОТРИСОВКИ =====
    private double[] displayedArray = new double[0];
    private int highlightedIndex1 = -1;
    private int highlightedIndex2 = -1;
    private string currentAlgorithmName = "";
    private double cachedMin = 0;
    private double cachedMax = 1;

    // ===== EXCEL =====
    private string lastExcelFilePath = "";
    private double[] lastSortedResult = null;

    private readonly SolidBrush normalBrush = new SolidBrush(Color.SteelBlue);
    private readonly SolidBrush highlightBrush1 = new SolidBrush(Color.Red);
    private readonly SolidBrush highlightBrush2 = new SolidBrush(Color.Orange);
    private readonly Font titleFont = new Font("Segoe UI", 14, FontStyle.Bold);
    private readonly Font infoFont = new Font("Segoe UI", 9);
    private readonly SolidBrush titleBrush = new SolidBrush(Color.DarkBlue);
    private readonly SolidBrush infoBrush = new SolidBrush(Color.Gray);

    public MainForm()
    {
      BuildUI();
    }

    // ===== СБОРКА ИНТЕРФЕЙСА =====
    private void BuildUI()
    {
      this.Text = "Сортировка";
      this.Size = new Size(1300, 880);
      this.StartPosition = FormStartPosition.CenterScreen;
      this.MinimumSize = new Size(1100, 750);

      var mainMenu = new MenuStrip();

      var fileMenu = new ToolStripMenuItem("Файл");
      fileMenu.DropDownItems.Add(new ToolStripMenuItem("Загрузить из Excel (.xlsx)", null, (s, e) => LoadFromExcel()));
      fileMenu.DropDownItems.Add(new ToolStripMenuItem("Загрузить из Google Table (CSV)", null, (s, e) => LoadFromGoogle()));
      fileMenu.DropDownItems.Add(new ToolStripSeparator());
      fileMenu.DropDownItems.Add(new ToolStripMenuItem("Сохранить отсортированное в Excel", null, (s, e) => SaveSortedToExcel()));
      fileMenu.DropDownItems.Add(new ToolStripSeparator());
      fileMenu.DropDownItems.Add(new ToolStripMenuItem("Выход", null, (s, e) => Close()));

      var actionsMenu = new ToolStripMenuItem("Действия");
      actionsMenu.DropDownItems.Add(new ToolStripMenuItem("Сгенерировать данные", null, (s, e) => GenerateData()));
      actionsMenu.DropDownItems.Add(new ToolStripSeparator());
      actionsMenu.DropDownItems.Add(new ToolStripMenuItem("Рассчитать", null, (s, e) => RunSorts()));
      actionsMenu.DropDownItems.Add(new ToolStripMenuItem("Очистить", null, (s, e) => ClearAll()));

      mainMenu.Items.Add(fileMenu);
      mainMenu.Items.Add(actionsMenu);
      this.MainMenuStrip = mainMenu;
      this.Controls.Add(mainMenu);

      statusLabel = new Label
      {
        Dock = DockStyle.Bottom,
        Height = 25,
        Text = "Готово",
        BackColor = Color.LightGray,
        TextAlign = ContentAlignment.MiddleLeft
      };
      this.Controls.Add(statusLabel);

      settingsPanel = new Panel
      {
        Dock = DockStyle.Left,
        Width = 340,
        Padding = new Padding(10),
        BackColor = Color.WhiteSmoke,
        AutoScroll = true
      };

      int y = 10;
      Action<string> addCaption = (text) => {
        settingsPanel.Controls.Add(new Label { Text = text, Location = new Point(10, y), AutoSize = true });
        y += 22;
      };

      addCaption("Количество чисел:");
      countTextBox = new TextBox { Location = new Point(10, y), Width = 300, Text = "20" };
      settingsPanel.Controls.Add(countTextBox);
      y += 32;

      settingsPanel.Controls.Add(new Label { Text = "Минимум:", Location = new Point(10, y), AutoSize = true });
      settingsPanel.Controls.Add(new Label { Text = "Максимум:", Location = new Point(165, y), AutoSize = true });
      y += 22;
      minValueTextBox = new TextBox { Location = new Point(10, y), Width = 140, Text = "-1000000000000" };
      maxValueTextBox = new TextBox { Location = new Point(165, y), Width = 145, Text = "1000000000000" };
      settingsPanel.Controls.Add(minValueTextBox);
      settingsPanel.Controls.Add(maxValueTextBox);
      y += 32;

      addCaption("Ручной ввод (числа через пробел, ';' или Enter):");
      manualInputTextBox = new TextBox
      {
        Location = new Point(10, y),
        Width = 300,
        Height = 22,
        Multiline = false,
        AcceptsReturn = false,
        AcceptsTab = false
      };
      manualInputTextBox.KeyDown += (s, e) => {
        if (e.KeyCode == Keys.Enter)
        {
          e.SuppressKeyPress = true;
          e.Handled = true;
          ApplyManualInput();
        }
      };
      settingsPanel.Controls.Add(manualInputTextBox);
      y += 30;

      applyManualButton = new Button
      {
        Text = "Применить ручной ввод",
        Location = new Point(10, y),
        Width = 300,
        Height = 30
      };
      applyManualButton.Click += (s, e) => ApplyManualInput();
      settingsPanel.Controls.Add(applyManualButton);
      y += 40;

      addCaption("Ограничение итераций (только для BOGO):");
      maxIterationsTextBox = new TextBox { Location = new Point(10, y), Width = 300, Text = "100000" };
      settingsPanel.Controls.Add(maxIterationsTextBox);
      y += 35;

      ascendingCheckBox = new CheckBox { Text = "По возрастанию", Location = new Point(10, y), Checked = true, AutoSize = true };
      settingsPanel.Controls.Add(ascendingCheckBox);
      y += 28;

      fractionalCheckBox = new CheckBox { Text = "Дробные числа", Location = new Point(10, y), Checked = true, AutoSize = true };
      settingsPanel.Controls.Add(fractionalCheckBox);
      y += 28;

      detailCheckBox = new CheckBox
      {
        Text = "Детализация (пошагово)",
        Location = new Point(10, y),
        Checked = true,
        AutoSize = true
      };
      detailCheckBox.CheckedChanged += (s, e) => ToggleVisualization();
      settingsPanel.Controls.Add(detailCheckBox);
      y += 32;

      addCaption("Выбор алгоритмов:");
      bubbleCheckBox = new CheckBox { Text = "Пузырьковая", Location = new Point(10, y), Checked = true, AutoSize = true }; y += 24;
      insertionCheckBox = new CheckBox { Text = "Вставками", Location = new Point(10, y), Checked = true, AutoSize = true }; y += 24;
      shakerCheckBox = new CheckBox { Text = "Шейкерная", Location = new Point(10, y), Checked = true, AutoSize = true }; y += 24;
      quickCheckBox = new CheckBox { Text = "Быстрая", Location = new Point(10, y), Checked = true, AutoSize = true }; y += 24;
      bogoCheckBox = new CheckBox { Text = "BOGO", Location = new Point(10, y), Checked = false, AutoSize = true };

      settingsPanel.Controls.AddRange(new Control[] { bubbleCheckBox, insertionCheckBox, shakerCheckBox, quickCheckBox, bogoCheckBox });

      this.Controls.Add(settingsPanel);

      contentPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(5) };

      var layout = new TableLayoutPanel
      {
        Dock = DockStyle.Fill,
        ColumnCount = 1,
        RowCount = 3
      };
      layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
      layout.RowStyles.Add(new RowStyle(SizeType.Percent, 40f));
      layout.RowStyles.Add(new RowStyle(SizeType.Percent, 25f));
      layout.RowStyles.Add(new RowStyle(SizeType.Percent, 35f));
      contentPanel.Controls.Add(layout);

      dataInputGrid = new DataGridView
      {
        Dock = DockStyle.Fill,
        AllowUserToAddRows = true,
        AllowUserToDeleteRows = true,
        AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
        RowHeadersVisible = false,
        BackgroundColor = Color.White,
        EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2,
        ScrollBars = ScrollBars.Both,
        AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None,
        AllowUserToResizeRows = false,
        ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
      };

      var valueColumn = new DataGridViewTextBoxColumn
      {
        Name = "ValueColumn",
        HeaderText = "Значение",
        AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
        SortMode = DataGridViewColumnSortMode.NotSortable
      };
      dataInputGrid.Columns.Add(valueColumn);

      layout.Controls.Add(dataInputGrid, 0, 0);

      visualizationPanel = new DoubleBufferedPanel
      {
        Dock = DockStyle.Fill,
        BackColor = Color.White,
        BorderStyle = BorderStyle.FixedSingle
      };
      visualizationPanel.Paint += VisualizationPanel_Paint;
      layout.Controls.Add(visualizationPanel, 0, 1);

      resultsGrid = new DataGridView
      {
        Dock = DockStyle.Fill,
        ReadOnly = true,
        AllowUserToAddRows = false,
        AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
        RowHeadersVisible = false,
        BackgroundColor = Color.White
      };
      resultsGrid.Columns.Add("AlgorithmColumn", "Алгоритм");
      resultsGrid.Columns.Add("TimeColumn", "Время (мс)");
      resultsGrid.Columns.Add("IterationsColumn", "Итераций (проходов)");
      resultsGrid.Columns.Add("StatusColumn", "Статус");
      layout.Controls.Add(resultsGrid, 0, 2);

      this.Controls.Add(contentPanel);
      contentPanel.BringToFront();
      settingsPanel.SendToBack();
    }

    private void ShowInputError(string message, string title = "Некорректный ввод")
    {
      MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }

    // ===== РУЧНОЙ ВВОД =====
    private void ApplyManualInput()
    {
      string manualRaw = manualInputTextBox.Text.Trim();
      if (string.IsNullOrEmpty(manualRaw))
      {
        ShowInputError("Поле ручного ввода пусто. Введите числа через пробел, например: 5 4 3 2 1");
        return;
      }

      if (manualRaw.Length > 2000000)
      {
        ShowInputError("Слишком много символов во вводе (максимум 2 миллиона).", "Ошибка");
        return;
      }

      var manualTokens = manualRaw.Split(new[] { ' ', '\t', '\n', '\r', ';' }, StringSplitOptions.RemoveEmptyEntries);
      var manualNumbers = new List<double>();

      if (manualTokens.Length > 0 && TryParseAll(manualTokens, manualNumbers))
      {
        CommitManual(manualNumbers);
        return;
      }

      manualNumbers.Clear();
      manualTokens = manualRaw.Split(new[] { ' ', '\t', '\n', '\r', ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
      if (manualTokens.Length > 0 && TryParseAll(manualTokens, manualNumbers))
      {
        CommitManual(manualNumbers);
        return;
      }

      ShowInputError(
          "Не удалось разобрать числа.\n\n" +
          "Примеры корректного ввода:\n" +
          "  5 4 3 2 1\n" +
          "  1.5 2.7 -3.14\n" +
          "  1,5 2,7 -3,14\n" +
          "  1.5; 2.7; -3.14",
          "Ошибка ручного ввода");
    }

    private void CommitManual(List<double> numbers)
    {
      SetData(numbers);
      manualInputTextBox.Clear();
      statusLabel.Text = $"Применён ручной ввод: загружено {numbers.Count} чисел";
    }

    private static bool TryParseNumber(string text, out double value)
    {
      value = 0;
      if (string.IsNullOrWhiteSpace(text)) return false;
      string normalized = text.Trim().Replace(',', '.');
      return double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
    }

    private static bool TryParseAll(string[] tokens, List<double> result)
    {
      foreach (var token in tokens)
      {
        if (!TryParseNumber(token, out double parsedValue)) return false;
        if (double.IsNaN(parsedValue) || double.IsInfinity(parsedValue)) return false;
        result.Add(parsedValue);
      }
      return true;
    }

    // ===== ЗАГРУЗКА ИЗ EXCEL =====
    private void LoadFromExcel()
    {
      using (var openDialog = new OpenFileDialog { Filter = "Excel|*.xlsx;*.xlsm|Все файлы|*.*" })
      {
        if (openDialog.ShowDialog() != DialogResult.OK) return;

        try
        {
          var parsedNumbers = new List<double>();
          using (var workbook = new XLWorkbook(openDialog.FileName))
          {
            var worksheet = workbook.Worksheet(1);
            var firstColumn = worksheet.Column(1);
            foreach (var cell in firstColumn.CellsUsed())
            {
              if (cell.DataType == XLDataType.Number)
              {
                parsedNumbers.Add(cell.GetDouble());
                continue;
              }
              string cellText = cell.GetString().Trim();
              if (string.IsNullOrEmpty(cellText)) continue;
              if (TryParseNumber(cellText, out double value))
                parsedNumbers.Add(value);
            }
          }

          if (parsedNumbers.Count == 0)
          {
            ShowInputError("В первом столбце Excel не найдено ни одного числа.", "Нет чисел");
            return;
          }

          lastExcelFilePath = openDialog.FileName;
          SetData(parsedNumbers);
          statusLabel.Text = $"Загружено {parsedNumbers.Count} чисел из Excel";
        }
        catch (FileNotFoundException) { ShowInputError("Файл не найден.", "Ошибка"); }
        catch (IOException) { ShowInputError("Не удалось открыть файл. Возможно, он занят Excel.", "Файл занят"); }
        catch (Exception exception) { ShowInputError("Ошибка загрузки Excel:\n\n" + exception.Message, "Ошибка"); }
      }
    }

    // ===== ЗАГРУЗКА ИЗ GOOGLE TABLE =====
    private void LoadFromGoogle()
    {
      string googleUrl = Microsoft.VisualBasic.Interaction.InputBox(
          "Как получить ссылку:\n" +
          "Google Таблица → Файл → Поделиться → Опубликовать в интернете →\n" +
          "вкладка «Ссылка» → выбрать лист → формат CSV → Опубликовать →\n" +
          "скопировать ссылку целиком (Ctrl+A → Ctrl+C).\n\n" +
          "Ссылка должна заканчиваться на ...output=csv\n\n" +
          "Вставьте ссылку:",
          "Загрузка из Google Table", "");

      if (string.IsNullOrWhiteSpace(googleUrl)) return;

      if (!googleUrl.StartsWith("http://") && !googleUrl.StartsWith("https://"))
      {
        ShowInputError("Ссылка должна начинаться с http:// или https://", "Неверная ссылка");
        return;
      }

      try
      {
        using (var webClient = new WebClient())
        {
          string downloadedContent = webClient.DownloadString(googleUrl);
          var parsedNumbers = new List<double>();

          foreach (var contentLine in downloadedContent.Split('\n'))
          {
            if (string.IsNullOrWhiteSpace(contentLine)) continue;
            var csvCells = ParseCsvLine(contentLine);
            if (csvCells.Count == 0) continue;

            string firstCell = csvCells[0].Trim();
            if (string.IsNullOrWhiteSpace(firstCell)) continue;

            if (TryParseNumber(firstCell, out double parsedValue))
              parsedNumbers.Add(parsedValue);
          }

          if (parsedNumbers.Count == 0)
          {
            ShowInputError(
                "По ссылке не найдено ни одного числа в первом столбце.\n\n" +
                "Проверьте:\n" +
                "  • Ссылка заканчивается на ...output=csv\n" +
                "  • В 1-м столбце таблицы есть числа",
                "Нет чисел");
            return;
          }

          SetData(parsedNumbers);
          statusLabel.Text = $"Загружено {parsedNumbers.Count} чисел из Google Table";
        }
      }
      catch (Exception exception)
      {
        ShowInputError("Ошибка загрузки Google Table:\n\n" + exception.Message, "Ошибка");
      }
    }

    private static List<string> ParseCsvLine(string line)
    {
      var result = new List<string>();
      var builder = new StringBuilder();
      bool inQuotes = false;

      for (int i = 0; i < line.Length; ++i)
      {
        char c = line[i];
        if (inQuotes)
        {
          if (c == '"')
          {
            if (i + 1 < line.Length && line[i + 1] == '"')
            {
              builder.Append('"');
              ++i;
            }
            else inQuotes = false;
          }
          else builder.Append(c);
        }
        else
        {
          if (c == '"') inQuotes = true;
          else if (c == ',') { result.Add(builder.ToString()); builder.Clear(); }
          else builder.Append(c);
        }
      }
      result.Add(builder.ToString());
      return result;
    }

    // ===== СОХРАНЕНИЕ В EXCEL =====
    private void SaveSortedToExcel()
    {
      if (string.IsNullOrEmpty(lastExcelFilePath) || !File.Exists(lastExcelFilePath))
      {
        MessageBox.Show("Excel-файл не загружен.\n\nПорядок: 1) Загрузить Excel, 2) Рассчитать, 3) Сохранить.",
            "Нет Excel-файла", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        return;
      }

      if (lastSortedResult == null || lastSortedResult.Length == 0)
      {
        MessageBox.Show("Нет отсортированных данных. Сначала: Действия → Рассчитать.",
            "Нет данных", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        return;
      }

      var sortedNumbers = lastSortedResult.ToList();

      var answer = MessageBox.Show(
          $"Отсортировано чисел: {sortedNumbers.Count}\n\nИсходный файл:\n{lastExcelFilePath}\n\n" +
          "Перезаписать исходный файл?\n  • Да — перезаписать\n  • Нет — сохранить как новый",
          "Сохранение в Excel", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
      if (answer == DialogResult.Cancel) return;

      string targetPath;
      if (answer == DialogResult.Yes)
        targetPath = lastExcelFilePath;
      else
      {
        using (var saveDialog = new SaveFileDialog
        {
          Filter = "Excel|*.xlsx",
          FileName = Path.GetFileNameWithoutExtension(lastExcelFilePath) + "_sorted.xlsx"
        })
        {
          if (saveDialog.ShowDialog() != DialogResult.OK) return;
          targetPath = saveDialog.FileName;
        }
      }

      try
      {
        if (answer == DialogResult.No)
          File.Copy(lastExcelFilePath, targetPath, overwrite: true);

        using (var workbook = new XLWorkbook(targetPath))
        {
          var worksheet = workbook.Worksheet(1);

          int startRow = 1;
          if (!TryParseNumber(worksheet.Cell(1, 1).GetString(), out _))
          {
            worksheet.Cell(1, 2).Value = "Отсортировано";
            startRow = 2;
          }

          for (int index = 0; index < sortedNumbers.Count; ++index)
            worksheet.Cell(startRow + index, 2).Value = sortedNumbers[index];

          workbook.Save();
        }

        statusLabel.Text = $"Сохранено {sortedNumbers.Count} чисел в Excel";
        MessageBox.Show($"Готово. Данные записаны во 2-й столбец:\n\n{targetPath}",
            "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
      }
      catch (IOException ioEx)
      {
        MessageBox.Show($"Не удалось сохранить: {ioEx.Message}\n\nЗакройте файл в Excel.",
            "Файл занят", MessageBoxButtons.OK, MessageBoxIcon.Error);
      }
      catch (Exception exception)
      {
        MessageBox.Show("Ошибка сохранения Excel:\n\n" + exception.Message, "Ошибка",
            MessageBoxButtons.OK, MessageBoxIcon.Error);
      }
    }

    // ===== ГЕНЕРАЦИЯ ДАННЫХ =====
    private void GenerateData()
    {
      if (!int.TryParse(countTextBox.Text.Trim(), out int requestedCount) || requestedCount <= 0 || requestedCount > 1000000)
      {
        ShowInputError("Количество должно быть целым числом от 1 до 1000000.", "Некорректное количество"); return;
      }
      if (!TryParseNumber(minValueTextBox.Text, out double minimumValue))
      {
        ShowInputError("Минимум должен быть числом.", "Некорректный минимум"); return;
      }
      if (!TryParseNumber(maxValueTextBox.Text, out double maximumValue))
      {
        ShowInputError("Максимум должен быть числом.", "Некорректный максимум"); return;
      }
      if (minimumValue >= maximumValue)
      {
        ShowInputError("Минимум должен быть меньше максимума.", "Неверный диапазон"); return;
      }

      if (minimumValue < -1e15 || maximumValue > 1e15)
      {
        ShowInputError("Диапазон не должен превышать ±10^15 (иначе теряется точность double).", "Слишком большой диапазон");
        return;
      }

      var randomGenerator = Random.Shared;
      var generatedNumbers = new List<double>(requestedCount);

      if (fractionalCheckBox.Checked)
      {
        double range = maximumValue - minimumValue;
        for (int counter = 0; counter < requestedCount; ++counter)
          generatedNumbers.Add(Math.Round(randomGenerator.NextDouble() * range + minimumValue, 4));
      }
      else
      {
        long intMin = (long)Math.Ceiling(minimumValue);
        long intMax = (long)Math.Floor(maximumValue);
        if (intMin > intMax) { ShowInputError("В диапазоне нет целых чисел.", "Ошибка"); return; }
        double intRange = (double)intMax - intMin + 1;
        for (int counter = 0; counter < requestedCount; ++counter)
          generatedNumbers.Add(intMin + Math.Floor(randomGenerator.NextDouble() * intRange));
      }

      SetData(generatedNumbers);
      statusLabel.Text = $"Сгенерировано {requestedCount} чисел";
    }

    // ===== УСТАНОВКА ДАННЫХ =====
    private void SetData(List<double> numbers)
    {
      currentData = new List<double>(numbers);

      while (dataInputGrid.Columns.Count > 1)
        dataInputGrid.Columns.RemoveAt(1);

      dataInputGrid.SuspendLayout();

      var oldAutoSize = dataInputGrid.AutoSizeColumnsMode;
      bool oldAllowAdd = dataInputGrid.AllowUserToAddRows;

      dataInputGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
      dataInputGrid.AllowUserToAddRows = false;

      dataInputGrid.Rows.Clear();
      if (numbers.Count > 0)
        dataInputGrid.RowCount = numbers.Count;

      for (int index = 0; index < numbers.Count; ++index)
        dataInputGrid.Rows[index].Cells[0].Value = numbers[index];

      dataInputGrid.AutoSizeColumnsMode = oldAutoSize;
      dataInputGrid.AllowUserToAddRows = oldAllowAdd;
      dataInputGrid.ResumeLayout();
      dataInputGrid.Refresh();
    }

    // ===== ЗАПИСЬ РЕЗУЛЬТАТА =====
    private void EnsureSortedColumns(List<SortBase> algorithms)
    {
      while (dataInputGrid.Columns.Count > 1)
        dataInputGrid.Columns.RemoveAt(1);

      foreach (var algorithm in algorithms)
      {
        dataInputGrid.Columns.Add(new DataGridViewTextBoxColumn
        {
          Name = "sorted_" + algorithm.Name,
          HeaderText = algorithm.Name,
          ReadOnly = true,
          SortMode = DataGridViewColumnSortMode.NotSortable,
          AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        });
      }
    }

    private void WriteSortedColumn(int columnIndex, double[] sortedArray)
    {
      dataInputGrid.SuspendLayout();

      var oldAutoSize = dataInputGrid.AutoSizeColumnsMode;
      bool oldAllowAdd = dataInputGrid.AllowUserToAddRows;

      dataInputGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
      dataInputGrid.AllowUserToAddRows = false;

      if (dataInputGrid.RowCount < sortedArray.Length)
        dataInputGrid.RowCount = sortedArray.Length;

      for (int rowIndex = 0; rowIndex < sortedArray.Length; ++rowIndex)
        dataInputGrid.Rows[rowIndex].Cells[columnIndex].Value = sortedArray[rowIndex];

      dataInputGrid.AutoSizeColumnsMode = oldAutoSize;
      dataInputGrid.AllowUserToAddRows = oldAllowAdd;
      dataInputGrid.ResumeLayout();
    }

    // ===== ГЛАВНЫЙ МЕТОД: ЗАПУСК ВСЕХ СОРТИРОВОК =====
    private async void RunSorts()
    {
      try
      {
        if (currentData == null || currentData.Count == 0)
        {
          ShowInputError("Нет данных для сортировки.", "Нет данных");
          return;
        }

        if (currentData.Count > DetailedThreshold && detailCheckBox.Checked)
        {
          var animAnswer = MessageBox.Show(
              $"Динамическая сортировка отображаться не будет — элементов больше {DetailedThreshold}.\n\n" +
              $"Показан будет только финальный отсортированный массив.\n\n" +
              $"Продолжить?",
              "Слишком много элементов для анимации",
              MessageBoxButtons.YesNo, MessageBoxIcon.Information);
          if (animAnswer == DialogResult.No) return;
        }

        if (currentData.Count > 50000)
        {
          var bigAnswer = MessageBox.Show(
              $"Сортировка {currentData.Count} элементов может занять несколько минут.\n\nПродолжить?",
              "Много элементов",
              MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
          if (bigAnswer == DialogResult.No) return;
        }

        long bogoLimit = long.MaxValue;
        if (bogoCheckBox.Checked)
        {
          if (!long.TryParse(maxIterationsTextBox.Text.Trim(), out bogoLimit) || bogoLimit <= 0)
          {
            ShowInputError("Лимит итераций — целое число > 0.", "Некорректный лимит");
            return;
          }
        }

        var selectedAlgorithms = new List<SortBase>();

        if (bubbleCheckBox.Checked) selectedAlgorithms.Add(new BubbleSort());
        if (insertionCheckBox.Checked) selectedAlgorithms.Add(new InsertionSort());
        if (shakerCheckBox.Checked) selectedAlgorithms.Add(new ShakerSort());
        if (quickCheckBox.Checked) selectedAlgorithms.Add(new QuickSort());
        if (bogoCheckBox.Checked) selectedAlgorithms.Add(new BogoSort { MaxIterations = bogoLimit });

        if (selectedAlgorithms.Count == 0) { ShowInputError("Не выбран ни один алгоритм.", "Алгоритмы не выбраны"); return; }

        if (bogoCheckBox.Checked && currentData.Count > 500)
        {
          var answer = MessageBox.Show(
              $"BOGO на {currentData.Count} элементах практически не имеет шанса завершиться успешно.\n" +
              "Скорее всего будет достигнут лимит итераций.\n\nПродолжить?",
              "BOGO на больших данных", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
          if (answer == DialogResult.No) return;
        }

        resultsGrid.Rows.Clear();
        lastSortedResult = null;

        bool ascending = ascendingCheckBox.Checked;
        bool detailed = currentData.Count <= DetailedThreshold && detailedVisualization;
        double[] baseArray = currentData.ToArray();

        EnsureSortedColumns(selectedAlgorithms);

        for (int algorithmIndex = 0; algorithmIndex < selectedAlgorithms.Count; ++algorithmIndex)
        {
          var algorithm = selectedAlgorithms[algorithmIndex];
          currentAlgorithmName = algorithm.Name;
          statusLabel.Text = $"Выполняется: {algorithm.Name}...";

          List<SortStep> recordedSteps = detailed ? new List<SortStep>() : null;
          var inputArray = (double[])baseArray.Clone();

          var sortResult = await Task.Run(() => {
            Action<SortStep> recorder = null;
            if (detailed)
            {
              var localSteps = recordedSteps;
              recorder = (step) => { if (localSteps.Count < MaxRecordedSteps) localSteps.Add(step); };
            }
            return algorithm.Sort(inputArray, ascending, recorder);
          });

          string statusText = sortResult.LimitReached ? "Лимит итераций" : "Завершено";
          resultsGrid.Rows.Add(sortResult.Name, sortResult.ElapsedMs.ToString("F4"),
              sortResult.Iterations, statusText);

          WriteSortedColumn(algorithmIndex + 1, sortResult.Result);
          lastSortedResult = sortResult.Result;

          if (detailed && recordedSteps != null && recordedSteps.Count > 0)
          {
            int totalSteps = recordedSteps.Count;
            int frameSkip = Math.Max(1, totalSteps / 300);

            for (int stepIndex = 0; stepIndex < totalSteps; stepIndex += frameSkip)
            {
              var step = recordedSteps[stepIndex];
              DrawArray(step.Array, step.Index1, step.Index2);
              await Task.Delay(10);
            }
          }

          DrawArray(sortResult.Result, -1, -1);
          await Task.Delay(150);
        }

        statusLabel.Text = "Все сортировки завершены";
      }
      catch (Exception exception)
      {
        MessageBox.Show("Непредвиденная ошибка:\n\n" + exception.Message,
            "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
      }
    }

    // ===== ВИЗУАЛИЗАЦИЯ =====
    private void DrawArray(double[] arrayToDraw, int firstHighlight, int secondHighlight)
    {
      displayedArray = arrayToDraw;
      highlightedIndex1 = firstHighlight;
      highlightedIndex2 = secondHighlight;

      if (arrayToDraw != null && arrayToDraw.Length > 0)
      {
        double minimumValue = arrayToDraw[0];
        double maximumValue = arrayToDraw[0];
        for (int i = 1; i < arrayToDraw.Length; ++i)
        {
          if (arrayToDraw[i] < minimumValue) minimumValue = arrayToDraw[i];
          if (arrayToDraw[i] > maximumValue) maximumValue = arrayToDraw[i];
        }
        cachedMin = minimumValue;
        cachedMax = maximumValue;
      }

      visualizationPanel.Invalidate();
    }

    private void VisualizationPanel_Paint(object sender, PaintEventArgs paintArgs)
    {
      if (displayedArray == null || displayedArray.Length == 0) return;

      var graphics = paintArgs.Graphics;
      graphics.Clear(visualizationPanel.BackColor);
      graphics.SmoothingMode = SmoothingMode.None;

      int panelWidth = visualizationPanel.Width;
      int panelHeight = visualizationPanel.Height;
      int elementCount = displayedArray.Length;

      if (elementCount > HistogramLimit)
      {
        string message = $"Слишком много элементов для отображения гистограммы ({elementCount})";
        var textSize = graphics.MeasureString(message, infoFont);
        graphics.DrawString(message, infoFont, infoBrush,
            (panelWidth - textSize.Width) / 2,
            (panelHeight - textSize.Height) / 2);
        return;
      }

      int step = Math.Max(1, elementCount / MaxRenderedBars);
      int shownCount = (elementCount + step - 1) / step;
      int barWidth = Math.Max(1, panelWidth / shownCount);

      double valueRange = Math.Max(0.0001, cachedMax - cachedMin);

      int drawIndex = 0;
      for (int elementIndex = 0; elementIndex < elementCount; elementIndex += step)
      {
        int barHeight = (int)((displayedArray[elementIndex] - cachedMin) / valueRange * (panelHeight - 30)) + 5;
        int barX = drawIndex * barWidth;
        int barY = panelHeight - barHeight;

        var barBrush = normalBrush;
        if (elementIndex == highlightedIndex1) barBrush = highlightBrush1;
        else if (elementIndex == highlightedIndex2) barBrush = highlightBrush2;

        graphics.FillRectangle(barBrush, barX, barY, Math.Max(1, barWidth - 1), barHeight);

        ++drawIndex;
      }

      graphics.DrawString(currentAlgorithmName, titleFont, titleBrush, 10, 8);

      string infoText = $"Элементов: {elementCount}";
      var infoSize = graphics.MeasureString(infoText, infoFont);
      graphics.DrawString(infoText, infoFont, infoBrush, panelWidth - infoSize.Width - 10, 12);
    }

    private void ToggleVisualization()
    {
      detailedVisualization = detailCheckBox.Checked;
      statusLabel.Text = detailedVisualization
          ? "Детализация: включена (мало чисел — по шагам)"
          : "Детализация: выключена (только финал)";
    }

    private void ClearAll()
    {
      currentData.Clear();
      while (dataInputGrid.Columns.Count > 1)
        dataInputGrid.Columns.RemoveAt(1);
      dataInputGrid.Rows.Clear();
      resultsGrid.Rows.Clear();
      displayedArray = new double[0];
      currentAlgorithmName = "";
      lastExcelFilePath = "";
      lastSortedResult = null;
      visualizationPanel.Invalidate();
      statusLabel.Text = "Очищено";
    }
  }
}