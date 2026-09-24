using System.Net;
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
              throw new Exception("Лист пуст");

            foreach (var cell in usedRange.CellsUsed())
            {
              if (double.TryParse(cell.GetString().Replace('.', ','), out double value))
                parsedNumbers.Add(value);
            }
          }
          LoadDataToGrid(parsedNumbers);
          statusLabel.Text = $"Загружено {parsedNumbers.Count} чисел из Excel";
        }
        catch (Exception exception)
        {
          MessageBox.Show("Ошибка загрузки Excel: " + exception.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
          LoadDataToGrid(parsedNumbers);
          statusLabel.Text = $"Загружено {parsedNumbers.Count} чисел из CSV";
        }
        catch (Exception exception)
        {
          MessageBox.Show("Ошибка загрузки CSV: " + exception.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
          LoadDataToGrid(parsedNumbers);
          statusLabel.Text = $"Загружено {parsedNumbers.Count} чисел из Google Table";
        }
      }
      catch (Exception exception)
      {
        MessageBox.Show("Ошибка загрузки Google Table: " + exception.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
      }
    }

    private void GenerateData()
    {
      try
      {
        if (!int.TryParse(countTextBox.Text, out int requestedCount))
          throw new Exception("Некорректное количество чисел");
        if (!double.TryParse(minValueTextBox.Text.Replace('.', ','), out double minimumValue))
          throw new Exception("Некорректный минимум");
        if (!double.TryParse(maxValueTextBox.Text.Replace('.', ','), out double maximumValue))
          throw new Exception("Некорректный максимум");

        if (requestedCount <= 0 || requestedCount > 100000)
          throw new Exception("Количество должно быть от 1 до 100000");
        if (minimumValue >= maximumValue)
          throw new Exception("Минимум должен быть меньше максимума");

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
          if (intMin > intMax)
            throw new Exception("В заданном диапазоне нет целых чисел");

          for (int counter = 0; counter < requestedCount; counter++)
            generatedNumbers.Add(randomGenerator.Next(intMin, intMax + 1));

          statusLabel.Text = $"Сгенерировано {requestedCount} целых чисел в [{intMin}; {intMax}]";
        }

        LoadDataToGrid(generatedNumbers);
      }
      catch (Exception exception)
      {
        MessageBox.Show("Ошибка генерации: " + exception.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
      }
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
        currentData = ReadDataFromGrid();

        if (currentData.Count == 0 && !string.IsNullOrWhiteSpace(manualInputTextBox.Text))
        {
          var manualTokens = manualInputTextBox.Text.Split(new[] { ' ', ',', ';', '\t' }, StringSplitOptions.RemoveEmptyEntries);
          var manualNumbers = new List<double>();
          foreach (var token in manualTokens)
          {
            if (double.TryParse(token.Replace('.', ','), out double parsedValue))
              manualNumbers.Add(parsedValue);
            else
            {
              MessageBox.Show($"Некорректное значение: '{token}'", "Ошибка ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
              return;
            }
          }
          currentData = manualNumbers;
        }

        if (currentData.Count == 0)
        {
          MessageBox.Show("Нет данных для сортировки. Сгенерируйте данные или введите вручную.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
          return;
        }

        bool ascending = ascendingCheckBox.Checked;
        long iterationLimit = long.MaxValue;
        if (!string.IsNullOrWhiteSpace(maxIterationsTextBox.Text))
        {
          if (!long.TryParse(maxIterationsTextBox.Text, out iterationLimit) || iterationLimit <= 0)
          {
            MessageBox.Show("Некорректное ограничение итераций", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
          MessageBox.Show("Выберите хотя бы один алгоритм", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
          return;
        }

        resultsGrid.Rows.Clear();

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
                {
                  this.BeginInvoke(new Action(() => DrawArray(step.Array, step.Index1, step.Index2)));
                }
                Thread.Sleep(15);
              };
            }
            return algorithm.Sort(snapshot, ascending, stepHandler);
          });

          string statusText = sortResult.LimitReached ? "Лимит итераций" : "Завершено";
          resultsGrid.Rows.Add(sortResult.Name, sortResult.ElapsedMs.ToString("F4"), sortResult.Iterations, statusText);

          DrawArray(sortResult.Result, -1, -1);
          await Task.Delay(200);
        }

        statusLabel.Text = "Все сортировки завершены";
      }
      catch (Exception exception)
      {
        MessageBox.Show("Ошибка: " + exception.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
      }
    }

    private List<double> ReadDataFromGrid()
    {
      var numbersList = new List<double>();
      foreach (DataGridViewRow gridRow in dataInputGrid.Rows)
      {
        if (gridRow.IsNewRow) continue;
        var cellValue = gridRow.Cells["ValueColumn"].Value;
        if (cellValue == null) continue;
        if (double.TryParse(cellValue.ToString().Replace('.', ','), out double parsedNumber))
          numbersList.Add(parsedNumber);
        else
          throw new Exception($"Некорректное значение в таблице: '{cellValue}'");
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