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
  public class MainForm : Form
  {
    private DataGridView dataInputGrid;
    private DataGridView resultsGrid;
    private Panel visualizationPanel;
    private Panel contentPanel;
    private Panel settingsPanel;

    private CheckBox bubbleCheckBox;
    private CheckBox insertionCheckBox;
    private CheckBox shakerCheckBox;
    private CheckBox quickCheckBox;
    private CheckBox bogoCheckBox;
    private CheckBox ascendingCheckBox;
    private CheckBox noAnimationCheckBox;
    private CheckBox fractionalCheckBox;

    private TextBox countTextBox;
    private TextBox minValueTextBox;
    private TextBox maxValueTextBox;
    private TextBox manualInputTextBox;
    private TextBox maxIterationsTextBox;

    private Label statusLabel;

    private List<double> currentData = new List<double>();
    private bool detailedVisualization = true;
    private const int DetailedThreshold = 50;

    private double[] displayedArray = new double[0];
    private int highlightedIndex1 = -1;
    private int highlightedIndex2 = -1;

    public MainForm()
    {
      BuildUI();
    }

    private void BuildUI()
    {
      this.Text = "Лабораторная работа: Алгоритмы сортировки";
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
      fileMenu.DropDownItems.Add(new ToolStripMenuItem("Выход", null, (s, e) => Close()));

      var actionsMenu = new ToolStripMenuItem("Действия");
      actionsMenu.DropDownItems.Add(new ToolStripMenuItem("Сгенерировать данные", null, (s, e) => GenerateData()));
      actionsMenu.DropDownItems.Add(new ToolStripMenuItem("Рассчитать", null, (s, e) => RunSorts()));
      actionsMenu.DropDownItems.Add(new ToolStripMenuItem("Очистить", null, (s, e) => ClearAll()));
      actionsMenu.DropDownItems.Add(new ToolStripMenuItem("Переключить детализацию", null, (s, e) => ToggleVisualization()));

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

      addCaption("Ручной ввод (пробел/запятая/;):");
      manualInputTextBox = new TextBox { Location = new Point(10, y), Width = 300, Height = 55, Multiline = true };
      settingsPanel.Controls.Add(manualInputTextBox);
      y += 65;

      addCaption("Ограничение итераций (BOGO):");
      maxIterationsTextBox = new TextBox { Location = new Point(10, y), Width = 300, Text = "10000" };
      settingsPanel.Controls.Add(maxIterationsTextBox);
      y += 35;

      ascendingCheckBox = new CheckBox { Text = "По возрастанию", Location = new Point(10, y), Checked = true, AutoSize = true };
      settingsPanel.Controls.Add(ascendingCheckBox);
      y += 28;

      noAnimationCheckBox = new CheckBox { Text = "Без анимации (реальное время)", Location = new Point(10, y), AutoSize = true };
      settingsPanel.Controls.Add(noAnimationCheckBox);
      y += 28;

      fractionalCheckBox = new CheckBox { Text = "Дробные числа", Location = new Point(10, y), Checked = true, AutoSize = true };
      settingsPanel.Controls.Add(fractionalCheckBox);
      y += 32;

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
        EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2
      };
      var valueColumn = new DataGridViewTextBoxColumn
      {
        Name = "ValueColumn",
        HeaderText = "Значение",
        AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
      };
      dataInputGrid.Columns.Add(valueColumn);
      layout.Controls.Add(dataInputGrid, 0, 0);

      visualizationPanel = new Panel
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

    // ============ Вспомогательные ============

    private void ShowInputError(string message, string title = "Некорректный ввод")
    {
      MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }

    // ============ Загрузка ============

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
            var usedRange = worksheet.RangeUsed();
            if (usedRange == null)
            {
              ShowInputError("Лист Excel пуст — нет данных для чтения.", "Пустой файл");
              return;
            }

            foreach (var cell in usedRange.CellsUsed())
            {
              if (double.TryParse(cell.GetString().Replace('.', ','), out double value))
                parsedNumbers.Add(value);
            }
          }

          if (parsedNumbers.Count == 0)
          {
            ShowInputError(
                "В файле Excel не найдено ни одного числа.\n\n" +
                "Проверьте, что первый лист содержит числовые значения.",
                "Нет чисел");
            return;
          }

          LoadDataToGrid(parsedNumbers);
          statusLabel.Text = $"Загружено {parsedNumbers.Count} чисел из Excel";
        }
        catch (FileNotFoundException)
        {
          ShowInputError("Файл не найден.", "Ошибка");
        }
        catch (IOException)
        {
          ShowInputError("Не удалось открыть файл.\n\nВозможно, он занят другой программой (Excel).", "Файл занят");
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
          if (fileLines.Length == 0)
          {
            ShowInputError("Файл пуст.", "Ошибка загрузки CSV");
            return;
          }

          var parsedNumbers = new List<double>();
          int skipped = 0;

          foreach (var fileLine in fileLines)
          {
            var tokens = fileLine.Split(new[] { ' ', ',', ';', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var token in tokens)
            {
              if (double.TryParse(token.Replace('.', ','), out double parsedValue))
                parsedNumbers.Add(parsedValue);
              else
                skipped++;
            }
          }

          if (parsedNumbers.Count == 0)
          {
            ShowInputError(
                "В файле не найдено ни одного числа.\n\n" +
                "Проверьте, что файл содержит числа (например: 5, -3.14, 0.5), " +
                "разделённые пробелом, запятой или точкой с запятой.",
                "Нет чисел в файле");
            return;
          }

          LoadDataToGrid(parsedNumbers);
          statusLabel.Text = $"Загружено {parsedNumbers.Count} чисел из CSV" +
                             (skipped > 0 ? $" (пропущено {skipped} нечисловых строк)" : "");
        }
        catch (IOException)
        {
          ShowInputError(
              "Не удалось прочитать файл.\n\n" +
              "Возможно, он открыт в другой программе или нет прав доступа.",
              "Ошибка чтения файла");
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
          "Введите ссылку на CSV-экспорт Google Table:\n\n" +
          "(Файл → Поделиться → Опубликовать в интернете → CSV)", "Google Table", "");

      if (string.IsNullOrWhiteSpace(googleUrl))
      {
        MessageBox.Show("Ссылка не введена. Операция отменена.", "Отмена",
            MessageBoxButtons.OK, MessageBoxIcon.Information);
        return;
      }

      if (!googleUrl.StartsWith("http://") && !googleUrl.StartsWith("https://"))
      {
        ShowInputError(
            "Некорректная ссылка.\n\nСсылка должна начинаться с http:// или https://\n" +
            $"Введено: \"{googleUrl}\"",
            "Неверная ссылка");
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
            var tokens = contentLine.Split(new[] { ' ', ',', ';', '\t', '\r' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var token in tokens)
            {
              if (double.TryParse(token.Replace('.', ','), out double parsedValue))
                parsedNumbers.Add(parsedValue);
            }
          }

          if (parsedNumbers.Count == 0)
          {
            ShowInputError(
                "По ссылке не найдено ни одного числа.\n\n" +
                "Проверьте, что таблица опубликована как CSV и содержит числа.",
                "Нет чисел");
            return;
          }

          LoadDataToGrid(parsedNumbers);
          statusLabel.Text = $"Загружено {parsedNumbers.Count} чисел из Google Table";
        }
      }
      catch (WebException webEx)
      {
        ShowInputError(
            "Не удалось скачать таблицу.\n\n" +
            $"Причина: {webEx.Message}\n\n" +
            "Проверьте интернет, правильность ссылки и доступ к таблице.",
            "Ошибка сети");
      }
      catch (Exception exception)
      {
        ShowInputError("Ошибка загрузки Google Table:\n\n" + exception.Message, "Ошибка");
      }
    }

    // ============ Генерация ============

    private void GenerateData()
    {
      // --- Количество ---
      string countRaw = countTextBox.Text.Trim();
      if (string.IsNullOrEmpty(countRaw))
      {
        ShowInputError("Количество не введено. Укажите целое число от 1 до 100000.");
        return;
      }
      if (!int.TryParse(countRaw, out int requestedCount))
      {
        ShowInputError($"Количество должно быть целым числом.\n\nВведено: \"{countRaw}\"\nПример: 20");
        return;
      }
      if (requestedCount <= 0)
      {
        ShowInputError($"Количество должно быть больше 0.\n\nВведено: {requestedCount}");
        return;
      }
      if (requestedCount > 100000)
      {
        ShowInputError($"Слишком много чисел.\n\nМаксимум: 100000\nВведено: {requestedCount}");
        return;
      }

      // --- Минимум ---
      string minRaw = minValueTextBox.Text.Trim().Replace('.', ',');
      if (string.IsNullOrEmpty(minRaw))
      {
        ShowInputError("Минимум не введён. Укажите число, например: -1 или 0,5");
        return;
      }
      if (!double.TryParse(minRaw, out double minimumValue))
      {
        ShowInputError($"Минимум должен быть числом.\n\nВведено: \"{minValueTextBox.Text}\"\nПример: -1 или 0,5");
        return;
      }

      // --- Максимум ---
      string maxRaw = maxValueTextBox.Text.Trim().Replace('.', ',');
      if (string.IsNullOrEmpty(maxRaw))
      {
        ShowInputError("Максимум не введён. Укажите число, например: 1 или 100");
        return;
      }
      if (!double.TryParse(maxRaw, out double maximumValue))
      {
        ShowInputError($"Максимум должен быть числом.\n\nВведено: \"{maxValueTextBox.Text}\"\nПример: 1 или 100");
        return;
      }

      // --- Логика диапазона ---
      if (minimumValue >= maximumValue)
      {
        ShowInputError($"Минимум должен быть меньше максимума.\n\nМинимум: {minimumValue}\nМаксимум: {maximumValue}");
        return;
      }

      // --- Проверка для целых ---
      if (!fractionalCheckBox.Checked)
      {
        int intMinCheck = (int)Math.Ceiling(minimumValue);
        int intMaxCheck = (int)Math.Floor(maximumValue);
        if (intMinCheck > intMaxCheck)
        {
          ShowInputError(
              $"В диапазоне [{minimumValue}; {maximumValue}] нет целых чисел.\n\n" +
              "Включите «Дробные числа» или расширьте диапазон.");
          return;
        }
      }

      // --- Генерация ---
      var randomGenerator = new Random();
      var generatedNumbers = new List<double>();

      if (fractionalCheckBox.Checked)
      {
        for (int counter = 0; counter < requestedCount; counter++)
        {
          double value = randomGenerator.NextDouble() * (maximumValue - minimumValue) + minimumValue;
          generatedNumbers.Add(Math.Round(value, 4));
        }
        statusLabel.Text = $"Сгенерировано {requestedCount} дробных чисел в [{minimumValue}; {maximumValue}]";
      }
      else
      {
        int intMin = (int)Math.Ceiling(minimumValue);
        int intMax = (int)Math.Floor(maximumValue);
        for (int counter = 0; counter < requestedCount; counter++)
          generatedNumbers.Add(randomGenerator.Next(intMin, intMax + 1));
        statusLabel.Text = $"Сгенерировано {requestedCount} целых чисел в [{intMin}; {intMax}]";
      }

      LoadDataToGrid(generatedNumbers);
    }

    private void LoadDataToGrid(List<double> numbers)
    {
      dataInputGrid.Rows.Clear();
      foreach (var number in numbers)
      {
        int rowIndex = dataInputGrid.Rows.Add();
        dataInputGrid.Rows[rowIndex].Cells["ValueColumn"].Value = number;
      }
      dataInputGrid.Refresh();
    }

    // ============ Запуск ============

    private async void RunSorts()
    {
      try
      {
        // ===== 1. Чтение данных из таблицы =====
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

        // ===== 2. Если пусто — пробуем ручной ввод =====
        if (currentData.Count == 0)
        {
          string manualRaw = manualInputTextBox.Text.Trim();
          if (string.IsNullOrEmpty(manualRaw))
          {
            ShowInputError(
                "Нет данных для сортировки.\n\n" +
                "Варианты:\n" +
                "  • Действия → Сгенерировать данные\n" +
                "  • Заполнить таблицу вручную\n" +
                "  • Ввести числа в поле «Ручной ввод»",
                "Нет данных");
            return;
          }

          var manualTokens = manualRaw.Split(new[] { ' ', ',', ';', '\t' }, StringSplitOptions.RemoveEmptyEntries);
          var manualNumbers = new List<double>();

          for (int tokenIndex = 0; tokenIndex < manualTokens.Length; tokenIndex++)
          {
            string token = manualTokens[tokenIndex].Replace('.', ',');
            if (!double.TryParse(token, out double parsedValue))
            {
              ShowInputError(
                  $"Некорректное число в позиции {tokenIndex + 1}: \"{manualTokens[tokenIndex]}\"\n\n" +
                  "Допустимы: 5, -3.14, 0,5. Уберите буквы и лишние символы.",
                  "Ошибка ручного ввода");
              return;
            }
            manualNumbers.Add(parsedValue);
          }
          currentData = manualNumbers;
        }

        if (currentData.Count == 0)
        {
          ShowInputError("После разбора данных список пуст.", "Нет данных");
          return;
        }

        // ===== 3. Лимит итераций =====
        long iterationLimit = long.MaxValue;
        string limitRaw = maxIterationsTextBox.Text.Trim();
        if (!string.IsNullOrEmpty(limitRaw))
        {
          if (!long.TryParse(limitRaw, out iterationLimit))
          {
            ShowInputError(
                $"Лимит итераций должен быть целым числом.\n\n" +
                $"Введено: \"{limitRaw}\"\nПример: 10000",
                "Некорректный лимит итераций");
            return;
          }
          if (iterationLimit <= 0)
          {
            ShowInputError(
                $"Лимит итераций должен быть больше 0.\n\n" +
                $"Введено: {iterationLimit}",
                "Некорректный лимит итераций");
            return;
          }
        }

        // ===== 4. Хотя бы один алгоритм =====
        var selectedAlgorithms = new List<SortBase>();
        if (bubbleCheckBox.Checked) selectedAlgorithms.Add(new BubbleSort { MaxIterations = iterationLimit });
        if (insertionCheckBox.Checked) selectedAlgorithms.Add(new InsertionSort { MaxIterations = iterationLimit });
        if (shakerCheckBox.Checked) selectedAlgorithms.Add(new ShakerSort { MaxIterations = iterationLimit });
        if (quickCheckBox.Checked) selectedAlgorithms.Add(new QuickSort { MaxIterations = iterationLimit });
        if (bogoCheckBox.Checked) selectedAlgorithms.Add(new BogoSort { MaxIterations = iterationLimit });

        if (selectedAlgorithms.Count == 0)
        {
          ShowInputError(
              "Не выбран ни один алгоритм.\n\n" +
              "Отметьте галочками нужные сортировки в разделе «Выбор алгоритмов» (слева). " +
              "Можно выбрать несколько.",
              "Алгоритмы не выбраны");
          return;
        }

        // ===== 5. Предупреждение о BOGO =====
        if (bogoCheckBox.Checked && currentData.Count > 10 && iterationLimit == long.MaxValue)
        {
          var answer = MessageBox.Show(
              $"BOGO на {currentData.Count} элементах без ограничения итераций " +
              "может работать практически вечно.\n\n" +
              "Продолжить без ограничения?",
              "BOGO без лимита",
              MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
          if (answer == DialogResult.No) return;
        }

        // ===== 6. Запуск =====
        resultsGrid.Rows.Clear();

        bool ascending = ascendingCheckBox.Checked;
        bool detailed = currentData.Count <= DetailedThreshold
                        && detailedVisualization
                        && !noAnimationCheckBox.Checked;

        foreach (var algorithm in selectedAlgorithms)
        {
          statusLabel.Text = $"Выполняется: {algorithm.Name}...";
          Application.DoEvents();

          double[] snapshot = currentData.ToArray();

          var sortResult = await Task.Run(() =>
          {
            Action<SortStep> stepHandler = null;
            if (detailed)
            {
              stepHandler = (step) =>
              {
                if (this.IsHandleCreated && !this.IsDisposed)
                  this.BeginInvoke(new Action(() => DrawArray(step.Array, step.Index1, step.Index2)));
                Thread.Sleep(15);
              };
            }
            return algorithm.Sort(snapshot, ascending, stepHandler);
          });

          string statusText = sortResult.LimitReached ? "Лимит итераций" : "Завершено";
          resultsGrid.Rows.Add(sortResult.Name, sortResult.ElapsedMs.ToString("F4"),
              sortResult.Iterations, statusText);

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
      int rowNumber = 0;

      foreach (DataGridViewRow gridRow in dataInputGrid.Rows)
      {
        if (gridRow.IsNewRow) continue;
        rowNumber++;

        var cellValue = gridRow.Cells["ValueColumn"].Value;

        if (cellValue == null || string.IsNullOrWhiteSpace(cellValue.ToString()))
          throw new Exception($"Строка {rowNumber} пуста.\n\n" +
                              "Удалите пустые строки в таблице или заполните их числами.");

        string text = cellValue.ToString().Trim().Replace('.', ',');

        if (!double.TryParse(text, out double parsedNumber))
          throw new Exception($"Строка {rowNumber}: \"{cellValue}\" — не число.\n\n" +
                              "Допустимы числа: 5, -3.14, 0,5. Буквы и посторонние символы запрещены.");

        if (double.IsNaN(parsedNumber) || double.IsInfinity(parsedNumber))
          throw new Exception($"Строка {rowNumber}: значение вне допустимого диапазона ({cellValue}).");

        numbersList.Add(parsedNumber);
      }

      return numbersList;
    }

    // ============ Визуализация ============

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
        {
          graphics.FillRectangle(brush, barX, barY, Math.Max(1, barWidth - 1), barHeight);
        }
      }

      graphics.DrawString($"Элементов: {elementCount}", this.Font, Brushes.Black, 5, 5);
    }

    private void ToggleVisualization()
    {
      detailedVisualization = !detailedVisualization;
      statusLabel.Text = detailedVisualization ? "Детализация: включена" : "Детализация: выключена";
    }

    private void ClearAll()
    {
      dataInputGrid.Rows.Clear();
      resultsGrid.Rows.Clear();
      displayedArray = new double[0];
      visualizationPanel.Invalidate();
      statusLabel.Text = "Очищено";
    }
  }
}