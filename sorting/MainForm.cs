using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace SortingLab
{
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
    private DataGridView dataInputGrid;
    private DataGridView resultsGrid;
    private DoubleBufferedPanel visualizationPanel;
    private Panel contentPanel;
    private Panel settingsPanel;

    private CheckBox bubbleCheckBox;
    private CheckBox insertionCheckBox;
    private CheckBox shakerCheckBox;
    private CheckBox quickCheckBox;
    private CheckBox bogoCheckBox;
    private CheckBox ascendingCheckBox;
    private CheckBox fractionalCheckBox;

    private TextBox countTextBox;
    private TextBox minValueTextBox;
    private TextBox maxValueTextBox;
    private TextBox manualInputTextBox;
    private TextBox maxIterationsTextBox;

    private Button applyManualButton;
    private Button toggleDetailButton;
    private Label statusLabel;

    private List<double> currentData = new List<double>();
    private bool detailedVisualization = true;
    private const int DetailedThreshold = 50;

    private double[] displayedArray = new double[0];
    private int highlightedIndex1 = -1;
    private int highlightedIndex2 = -1;
    private string currentAlgorithmName = "";

    private string lastExcelFilePath = "";
    private double[] lastSortedResult = null;

    public MainForm()
    {
      BuildUI();
    }

    private void BuildUI()
    {
      this.Text = "Алгоритмы сортировки";
      this.Size = new Size(1300, 880);
      this.StartPosition = FormStartPosition.CenterScreen;
      this.MinimumSize = new Size(1100, 750);

      // ==== MenuStrip ====
      var mainMenu = new MenuStrip();

      var fileMenu = new ToolStripMenuItem("Файл");
      fileMenu.DropDownItems.Add(new ToolStripMenuItem("Загрузить из Excel (.xlsx)", null, (s, e) => LoadFromExcel()));
      fileMenu.DropDownItems.Add(new ToolStripMenuItem("Загрузить из CSV/TXT", null, (s, e) => LoadFromCsv()));
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

      // ==== Status ====
      statusLabel = new Label
      {
        Dock = DockStyle.Bottom,
        Height = 25,
        Text = "Готово",
        BackColor = Color.LightGray,
        TextAlign = ContentAlignment.MiddleLeft
      };
      this.Controls.Add(statusLabel);

      // ==== Left panel ====
      settingsPanel = new Panel
      {
        Dock = DockStyle.Left,
        Width = 340,
        Padding = new Padding(10),
        BackColor = Color.WhiteSmoke,
        AutoScroll = true
      };

      int y = 10;
      Action<string> addCaption = (text) =>
      {
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
      minValueTextBox = new TextBox { Location = new Point(10, y), Width = 140, Text = "-1" };
      maxValueTextBox = new TextBox { Location = new Point(165, y), Width = 145, Text = "1" };
      settingsPanel.Controls.Add(minValueTextBox);
      settingsPanel.Controls.Add(maxValueTextBox);
      y += 32;

      addCaption("Ручной ввод (числа только через ПРОБЕЛ):");
      manualInputTextBox = new TextBox
      {
        Location = new Point(10, y),
        Width = 300,
        Height = 22,
        Multiline = false,
        AcceptsReturn = false,
        AcceptsTab = false
      };
      manualInputTextBox.KeyDown += (s, e) =>
      {
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

      addCaption("Ограничение итераций (BOGO):");
      maxIterationsTextBox = new TextBox { Location = new Point(10, y), Width = 300, Text = "10000" };
      settingsPanel.Controls.Add(maxIterationsTextBox);
      y += 35;

      ascendingCheckBox = new CheckBox { Text = "По возрастанию", Location = new Point(10, y), Checked = true, AutoSize = true };
      settingsPanel.Controls.Add(ascendingCheckBox);
      y += 28;

      fractionalCheckBox = new CheckBox { Text = "Дробные числа", Location = new Point(10, y), Checked = true, AutoSize = true };
      settingsPanel.Controls.Add(fractionalCheckBox);
      y += 28;

      toggleDetailButton = new Button
      {
        Text = "Детализация: включена",
        Location = new Point(10, y),
        Width = 300,
        Height = 28
      };
      toggleDetailButton.Click += (s, e) => ToggleVisualization();
      settingsPanel.Controls.Add(toggleDetailButton);
      y += 38;

      addCaption("Выбор алгоритмов:");
      bubbleCheckBox = new CheckBox { Text = "Пузырьковая", Location = new Point(10, y), Checked = true, AutoSize = true }; y += 24;
      insertionCheckBox = new CheckBox { Text = "Вставками", Location = new Point(10, y), Checked = true, AutoSize = true }; y += 24;
      shakerCheckBox = new CheckBox { Text = "Шейкерная", Location = new Point(10, y), Checked = true, AutoSize = true }; y += 24;
      quickCheckBox = new CheckBox { Text = "Быстрая", Location = new Point(10, y), Checked = true, AutoSize = true }; y += 24;
      bogoCheckBox = new CheckBox { Text = "BOGO", Location = new Point(10, y), Checked = false, AutoSize = true };

      settingsPanel.Controls.AddRange(new Control[] { bubbleCheckBox, insertionCheckBox, shakerCheckBox, quickCheckBox, bogoCheckBox });

      this.Controls.Add(settingsPanel);

      // ==== Right (Fill) ====
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
        AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
      };
      var sortedColumn = new DataGridViewTextBoxColumn
      {
        Name = "SortedColumn",
        HeaderText = "Отсортировано",
        AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
        ReadOnly = true
      };
      dataInputGrid.Columns.Add(valueColumn);
      dataInputGrid.Columns.Add(sortedColumn);

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
      resultsGrid.Columns.Add("IterationsColumn", "Итераций");
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

    private void ApplyManualInput()
    {
      string manualRaw = manualInputTextBox.Text.Trim();
      if (string.IsNullOrEmpty(manualRaw))
      {
        ShowInputError("Поле ручного ввода пусто. Введите числа через пробел, например: 5 4 3 2 1");
        return;
      }

      var manualTokens = manualRaw.Split(new[] { ' ', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
      var manualNumbers = new List<double>();

      for (int tokenIndex = 0; tokenIndex < manualTokens.Length; tokenIndex++)
      {
        string normalized = manualTokens[tokenIndex].Replace('.', ',');
        if (!double.TryParse(normalized, out double parsedValue))
        {
          ShowInputError(
              $"Некорректное число в позиции {tokenIndex + 1}: \"{manualTokens[tokenIndex]}\"\n\n" +
              "Числа разделяются ТОЛЬКО пробелом. Например: 5 4 3 2 1",
              "Ошибка ручного ввода");
          return;
        }
        manualNumbers.Add(parsedValue);
      }

      LoadDataToGrid(manualNumbers);
      manualInputTextBox.Clear();
      statusLabel.Text = $"Применён ручной ввод: загружено {manualNumbers.Count} чисел";
    }

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
              if (double.TryParse(cell.GetString().Replace('.', ','), out double value))
                parsedNumbers.Add(value);
            }
          }

          if (parsedNumbers.Count == 0)
          {
            ShowInputError("В первом столбце Excel не найдено ни одного числа.", "Нет чисел");
            return;
          }

          lastExcelFilePath = openDialog.FileName;
          LoadDataToGrid(parsedNumbers);
          statusLabel.Text = $"Загружено {parsedNumbers.Count} чисел из Excel";
        }
        catch (FileNotFoundException)
        {
          ShowInputError("Файл не найден.", "Ошибка");
        }
        catch (IOException)
        {
          ShowInputError("Не удалось открыть файл. Возможно, он занят Excel.", "Файл занят");
        }
        catch (Exception exception)
        {
          ShowInputError("Ошибка загрузки Excel:\n\n" + exception.Message, "Ошибка");
        }
      }
    }

    private void LoadFromCsv()
    {
      using (var openDialog = new OpenFileDialog { Filter = "CSV/TXT|*.csv;*.txt|Все файлы|*.*" })
      {
        if (openDialog.ShowDialog() != DialogResult.OK) return;
        try
        {
          var fileLines = File.ReadAllLines(openDialog.FileName);
          var parsedNumbers = new List<double>();
          foreach (var fileLine in fileLines)
          {
            var tokens = fileLine.Split(new[] { ' ', ',', ';', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var token in tokens)
            {
              if (double.TryParse(token.Replace('.', ','), out double parsedValue))
                parsedNumbers.Add(parsedValue);
            }
          }
          if (parsedNumbers.Count == 0)
          {
            ShowInputError("В файле не найдено ни одного числа.", "Нет чисел");
            return;
          }
          LoadDataToGrid(parsedNumbers);
          statusLabel.Text = $"Загружено {parsedNumbers.Count} чисел из CSV";
        }
        catch (Exception exception)
        {
          ShowInputError("Ошибка загрузки CSV:\n\n" + exception.Message, "Ошибка");
        }
      }
    }

    private void LoadFromGoogle()
    {
      string googleUrl = Microsoft.VisualBasic.Interaction.InputBox(
          "Введите ссылку на CSV-экспорт Google Table:", "Google Table", "");
      if (string.IsNullOrWhiteSpace(googleUrl)) return;
      try
      {
        using (var webClient = new WebClient())
        {
          string downloadedContent = webClient.DownloadString(googleUrl);
          var parsedNumbers = new List<double>();
          foreach (var contentLine in downloadedContent.Split('\n'))
          {
            var tokens = contentLine.Split(new[] { ' ', ',', ';', '\t', '\r' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var token in tokens)
            {
              if (double.TryParse(token.Replace('.', ','), out double parsedValue))
                parsedNumbers.Add(parsedValue);
            }
          }
          if (parsedNumbers.Count == 0)
          {
            ShowInputError("По ссылке не найдено ни одного числа.", "Нет чисел");
            return;
          }
          LoadDataToGrid(parsedNumbers);
          statusLabel.Text = $"Загружено {parsedNumbers.Count} чисел из Google Table";
        }
      }
      catch (Exception exception)
      {
        ShowInputError("Ошибка загрузки Google Table:\n\n" + exception.Message, "Ошибка");
      }
    }

    private void SaveSortedToExcel()
    {
      if (string.IsNullOrEmpty(lastExcelFilePath) || !File.Exists(lastExcelFilePath))
      {
        MessageBox.Show(
            "Excel-файл не загружен.\n\n" +
            "Порядок работы:\n" +
            "  1. Файл → Загрузить из Excel (.xlsx)\n" +
            "  2. Действия → Рассчитать\n" +
            "  3. Файл → Сохранить отсортированное в Excel",
            "Нет Excel-файла",
            MessageBoxButtons.OK, MessageBoxIcon.Warning);
        return;
      }

      var sortedNumbers = new List<double>();
      foreach (DataGridViewRow gridRow in dataInputGrid.Rows)
      {
        if (gridRow.IsNewRow) continue;
        var sortedCell = gridRow.Cells["SortedColumn"].Value;
        if (sortedCell != null && double.TryParse(sortedCell.ToString().Replace('.', ','), out double parsed))
          sortedNumbers.Add(parsed);
      }

      if (sortedNumbers.Count == 0 && lastSortedResult != null)
        sortedNumbers = lastSortedResult.ToList();

      if (sortedNumbers.Count == 0)
      {
        MessageBox.Show(
            "Нет отсортированных данных.\n\n" +
            "Сначала выполните сортировку: Действия → Рассчитать",
            "Нет данных",
            MessageBoxButtons.OK, MessageBoxIcon.Warning);
        return;
      }

      var answer = MessageBox.Show(
          $"Отсортировано чисел: {sortedNumbers.Count}\n\n" +
          $"Исходный файл:\n{lastExcelFilePath}\n\n" +
          "Перезаписать исходный файл?\n" +
          "  • Да — перезаписать (2-й столбец заполнится)\n" +
          "  • Нет — сохранить как новый файл",
          "Сохранение в Excel",
          MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

      if (answer == DialogResult.Cancel) return;

      string targetPath;
      if (answer == DialogResult.Yes)
      {
        targetPath = lastExcelFilePath;
      }
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
          var firstCell = worksheet.Cell(1, 1);
          if (!double.TryParse(firstCell.GetString().Replace('.', ','), out _))
          {
            worksheet.Cell(1, 2).Value = "Отсортировано";
            startRow = 2;
          }

          for (int index = 0; index < sortedNumbers.Count; index++)
          {
            worksheet.Cell(startRow + index, 2).Value = sortedNumbers[index];
          }

          workbook.Save();
        }

        statusLabel.Text = $"Сохранено {sortedNumbers.Count} чисел в Excel";
        MessageBox.Show(
            $"Готово. Отсортированные данные записаны во 2-й столбец:\n\n{targetPath}",
            "Сохранено",
            MessageBoxButtons.OK, MessageBoxIcon.Information);
      }
      catch (IOException ioEx)
      {
        MessageBox.Show(
            "Не удалось сохранить файл.\n\n" +
            $"Причина: {ioEx.Message}\n\n" +
            "Скорее всего файл открыт в Excel — закройте его.",
            "Файл занят",
            MessageBoxButtons.OK, MessageBoxIcon.Error);
      }
      catch (Exception exception)
      {
        MessageBox.Show(
            "Ошибка сохранения Excel:\n\n" + exception.Message,
            "Ошибка",
            MessageBoxButtons.OK, MessageBoxIcon.Error);
      }
    }

    private void GenerateData()
    {
      if (!int.TryParse(countTextBox.Text.Trim(), out int requestedCount) || requestedCount <= 0 || requestedCount > 100000)
      {
        ShowInputError("Количество должно быть целым числом от 1 до 100000.", "Некорректное количество");
        return;
      }
      if (!double.TryParse(minValueTextBox.Text.Trim().Replace('.', ','), out double minimumValue))
      {
        ShowInputError("Минимум должен быть числом.", "Некорректный минимум");
        return;
      }
      if (!double.TryParse(maxValueTextBox.Text.Trim().Replace('.', ','), out double maximumValue))
      {
        ShowInputError("Максимум должен быть числом.", "Некорректный максимум");
        return;
      }
      if (minimumValue >= maximumValue)
      {
        ShowInputError("Минимум должен быть меньше максимума.", "Неверный диапазон");
        return;
      }

      var randomGenerator = new Random();
      var generatedNumbers = new List<double>();

      if (fractionalCheckBox.Checked)
      {
        for (int counter = 0; counter < requestedCount; counter++)
          generatedNumbers.Add(Math.Round(randomGenerator.NextDouble() * (maximumValue - minimumValue) + minimumValue, 4));
      }
      else
      {
        int intMin = (int)Math.Ceiling(minimumValue);
        int intMax = (int)Math.Floor(maximumValue);
        if (intMin > intMax)
        {
          ShowInputError("В диапазоне нет целых чисел.", "Ошибка");
          return;
        }
        for (int counter = 0; counter < requestedCount; counter++)
          generatedNumbers.Add(randomGenerator.Next(intMin, intMax + 1));
      }

      LoadDataToGrid(generatedNumbers);
      statusLabel.Text = $"Сгенерировано {requestedCount} чисел";
    }

    private void LoadDataToGrid(List<double> numbers)
    {
      dataInputGrid.Rows.Clear();
      foreach (var number in numbers)
      {
        int rowIndex = dataInputGrid.Rows.Add();
        dataInputGrid.Rows[rowIndex].Cells["ValueColumn"].Value = number;
        dataInputGrid.Rows[rowIndex].Cells["SortedColumn"].Value = null;
      }
      dataInputGrid.Refresh();
    }

    private void WriteSortedToSecondColumn(double[] sortedArray)
    {
      while (dataInputGrid.Rows.Count < sortedArray.Length)
        dataInputGrid.Rows.Add();

      for (int rowIndex = 0; rowIndex < sortedArray.Length; rowIndex++)
        dataInputGrid.Rows[rowIndex].Cells["SortedColumn"].Value = sortedArray[rowIndex];

      for (int rowIndex = sortedArray.Length; rowIndex < dataInputGrid.Rows.Count; rowIndex++)
      {
        if (dataInputGrid.Rows[rowIndex].IsNewRow) continue;
        dataInputGrid.Rows[rowIndex].Cells["SortedColumn"].Value = null;
      }

      dataInputGrid.Refresh();
    }

    private async void RunSorts()
    {
      try
      {
        try
        {
          currentData = ReadDataFromGrid();
        }
        catch (Exception dataEx)
        {
          MessageBox.Show(dataEx.Message, "Ошибка в таблице данных",
              MessageBoxButtons.OK, MessageBoxIcon.Warning);
          return;
        }

        if (currentData.Count == 0)
        {
          ShowInputError("Нет данных для сортировки.", "Нет данных");
          return;
        }

        long iterationLimit = long.MaxValue;
        string limitRaw = maxIterationsTextBox.Text.Trim();
        if (!string.IsNullOrEmpty(limitRaw))
        {
          if (!long.TryParse(limitRaw, out iterationLimit) || iterationLimit <= 0)
          {
            ShowInputError("Лимит итераций — целое число > 0.", "Некорректный лимит");
            return;
          }
        }

        var selectedAlgorithms = new List<SortBase>();
        if (bubbleCheckBox.Checked) selectedAlgorithms.Add(new BubbleSort { MaxIterations = iterationLimit });
        if (insertionCheckBox.Checked) selectedAlgorithms.Add(new InsertionSort { MaxIterations = iterationLimit });
        if (shakerCheckBox.Checked) selectedAlgorithms.Add(new ShakerSort { MaxIterations = iterationLimit });
        if (quickCheckBox.Checked) selectedAlgorithms.Add(new QuickSort { MaxIterations = iterationLimit });
        if (bogoCheckBox.Checked) selectedAlgorithms.Add(new BogoSort { MaxIterations = iterationLimit });

        if (selectedAlgorithms.Count == 0)
        {
          ShowInputError("Не выбран ни один алгоритм.", "Алгоритмы не выбраны");
          return;
        }

        if (bogoCheckBox.Checked && currentData.Count > 10 && iterationLimit == long.MaxValue)
        {
          var answer = MessageBox.Show(
              $"BOGO на {currentData.Count} элементах без лимита может работать вечно. Продолжить?",
              "BOGO без лимита",
              MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
          if (answer == DialogResult.No) return;
        }

        resultsGrid.Rows.Clear();
        lastSortedResult = null;

        bool ascending = ascendingCheckBox.Checked;
        bool detailed = currentData.Count <= DetailedThreshold && detailedVisualization;

        foreach (var algorithm in selectedAlgorithms)
        {
          statusLabel.Text = $"Выполняется: {algorithm.Name}...";
          currentAlgorithmName = algorithm.Name;
          Application.DoEvents();

          double[] snapshot = currentData.ToArray();
          List<SortStep> recordedSteps = detailed ? new List<SortStep>() : null;

          var sortResult = await Task.Run(() =>
          {
            Action<SortStep> recorder = null;
            if (detailed)
              recorder = (step) => recordedSteps.Add(step);

            return algorithm.Sort(snapshot, ascending, recorder);
          });

          string statusText = sortResult.LimitReached ? "Лимит итераций" : "Завершено";
          resultsGrid.Rows.Add(sortResult.Name, sortResult.ElapsedMs.ToString("F4"),
              sortResult.Iterations, statusText);

          WriteSortedToSecondColumn(sortResult.Result);
          lastSortedResult = sortResult.Result;

          if (detailed && recordedSteps != null && recordedSteps.Count > 0)
          {
            int totalSteps = recordedSteps.Count;
            int frameSkip = Math.Max(1, totalSteps / 300);

            for (int stepIndex = 0; stepIndex < totalSteps; stepIndex += frameSkip)
            {
              var step = recordedSteps[stepIndex];
              DrawArray(step.Array, step.Index1, step.Index2);
              await Task.Delay(15);
            }
          }

          DrawArray(sortResult.Result, -1, -1);
          await Task.Delay(200);
        }

        statusLabel.Text = "Все сортировки завершены";
      }
      catch (Exception exception)
      {
        MessageBox.Show("Непредвиденная ошибка:\n\n" + exception.Message,
            "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
      }
    }

    private List<double> ReadDataFromGrid()
    {
      var numbersList = new List<double>();
      var emptyRows = new List<int>();
      int rowNumber = 0;

      foreach (DataGridViewRow gridRow in dataInputGrid.Rows)
      {
        if (gridRow.IsNewRow) continue;

        rowNumber++;
        var cellValue = gridRow.Cells["ValueColumn"].Value;

        if (cellValue == null || string.IsNullOrWhiteSpace(cellValue.ToString()))
        {
          emptyRows.Add(rowNumber);
          continue;
        }

        if (!double.TryParse(cellValue.ToString().Trim().Replace('.', ','), out double parsedNumber))
          throw new Exception($"Строка {rowNumber}: \"{cellValue}\" — не число.\n\n" +
                              "Допустимы: 5, -3.14, 0,5. Буквы и посторонние символы запрещены.");

        numbersList.Add(parsedNumber);
      }

      if (emptyRows.Count > 0 && numbersList.Count > 0)
      {
        bool emptyInMiddle = false;
        int lastNonEmptyIndex = -1;

        for (int index = 0; index < dataInputGrid.Rows.Count; index++)
        {
          if (dataInputGrid.Rows[index].IsNewRow) break;
          var value = dataInputGrid.Rows[index].Cells["ValueColumn"].Value;
          if (value != null && !string.IsNullOrWhiteSpace(value.ToString()))
            lastNonEmptyIndex = index;
        }

        for (int index = 0; index < lastNonEmptyIndex; index++)
        {
          var value = dataInputGrid.Rows[index].Cells["ValueColumn"].Value;
          if (value == null || string.IsNullOrWhiteSpace(value.ToString()))
          {
            emptyInMiddle = true;
            break;
          }
        }

        if (emptyInMiddle)
        {
          throw new Exception(
              "В таблице есть пустые строки между числами.\n\n" +
              $"Пустые строки: {string.Join(", ", emptyRows)}\n\n" +
              "Удалите пустые строки (выделите → Delete) или заполните их числами.");
        }
      }

      return numbersList;
    }

    private void DrawArray(double[] arrayToDraw, int firstHighlight, int secondHighlight)
    {
      displayedArray = arrayToDraw;
      highlightedIndex1 = firstHighlight;
      highlightedIndex2 = secondHighlight;
      visualizationPanel.Invalidate();
    }

    private void VisualizationPanel_Paint(object sender, PaintEventArgs paintArgs)
    {
      if (displayedArray == null || displayedArray.Length == 0) return;

      var graphics = paintArgs.Graphics;
      graphics.Clear(visualizationPanel.BackColor);
      graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

      int panelWidth = visualizationPanel.Width;
      int panelHeight = visualizationPanel.Height;
      int elementCount = displayedArray.Length;

      int barWidth = Math.Max(1, panelWidth / elementCount);
      double maximumValue = displayedArray.Max();
      double minimumValue = displayedArray.Min();
      double valueRange = Math.Max(0.0001, maximumValue - minimumValue);

      for (int elementIndex = 0; elementIndex < elementCount; elementIndex++)
      {
        int barHeight = (int)((displayedArray[elementIndex] - minimumValue) / valueRange * (panelHeight - 30)) + 5;
        int barX = elementIndex * barWidth;
        int barY = panelHeight - barHeight;

        Color barColor = Color.SteelBlue;
        if (elementIndex == highlightedIndex1) barColor = Color.Red;
        else if (elementIndex == highlightedIndex2) barColor = Color.Orange;

        using (var brush = new SolidBrush(barColor))
          graphics.FillRectangle(brush, barX, barY, Math.Max(1, barWidth - 1), barHeight);
      }

      using (var titleFont = new Font("Segoe UI", 14, FontStyle.Bold))
      using (var titleBrush = new SolidBrush(Color.DarkBlue))
        graphics.DrawString(currentAlgorithmName, titleFont, titleBrush, 10, 8);

      using (var infoFont = new Font("Segoe UI", 9))
      using (var infoBrush = new SolidBrush(Color.Gray))
      {
        string infoText = $"Элементов: {elementCount}";
        var textSize = graphics.MeasureString(infoText, infoFont);
        graphics.DrawString(infoText, infoFont, infoBrush, panelWidth - textSize.Width - 10, 12);
      }
    }

    private void ToggleVisualization()
    {
      detailedVisualization = !detailedVisualization;
      toggleDetailButton.Text = detailedVisualization ? "Детализация: включена" : "Детализация: выключена";
    }

    private void ClearAll()
    {
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