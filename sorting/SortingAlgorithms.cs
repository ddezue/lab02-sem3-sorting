using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace SortingLab
{
  // Один кадр визуализации
  public class SortStep
  {
    public double[] Array { get; set; }
    public int Index1 { get; set; } = -1;
    public int Index2 { get; set; } = -1;
    public string Comment { get; set; } = "";
  }

  // Итог работы алгоритма
  public class SortResult
  {
    public string Name { get; set; }
    public double ElapsedMs { get; set; }
    public long Iterations { get; set; }
    public bool LimitReached { get; set; }
    public double[] Result { get; set; }
  }

  // Базовый класс для всех сортировок
  public abstract class SortBase
  {
    public string Name { get; protected set; }
    public long Iterations { get; protected set; }

    public abstract SortResult Sort(double[] input, bool ascending, Action<SortStep> onStep = null);
    protected static double[] Copy(double[] source) => (double[])source.Clone();

    protected static void EmitStep(Action<SortStep> onStep, Stopwatch stopwatch, double[] array,
        int index1 = -1, int index2 = -1, string comment = "")
    {
      if (onStep == null)
        return;

      stopwatch.Stop();
      var step = new SortStep
      {
        Array = Copy(array),
        Index1 = index1,
        Index2 = index2,
        Comment = comment
      };
      onStep(step);
      stopwatch.Start();
    }
  }

  // ===== ПУЗЫРЬКОВАЯ ===== (итерация = проход, макс n-1)
  public class BubbleSort : SortBase
  {
    public BubbleSort()
    {
      Name = "Пузырьковая";
    }

    public override SortResult Sort(double[] input, bool ascending, Action<SortStep> onStep)
    {
      Iterations = 0;
      var workingArray = Copy(input);
      var stopwatch = Stopwatch.StartNew();
      int arrayLength = workingArray.Length;

      for (int outerPass = 0; outerPass < arrayLength - 1; ++outerPass)
      {
        ++Iterations;
        bool swapped = false;

        for (int innerIndex = 0; innerIndex < arrayLength - 1 - outerPass; ++innerIndex)
        {
          bool needSwap = ascending
              ? workingArray[innerIndex] > workingArray[innerIndex + 1]
              : workingArray[innerIndex] < workingArray[innerIndex + 1];

          if (needSwap)
          {
            (workingArray[innerIndex], workingArray[innerIndex + 1]) =
                (workingArray[innerIndex + 1], workingArray[innerIndex]);
            swapped = true;
          }
        }

        EmitStep(onStep, stopwatch, workingArray);

        if (!swapped) break;
      }

      stopwatch.Stop();
      return new SortResult
      {
        Name = Name,
        ElapsedMs = stopwatch.Elapsed.TotalMilliseconds,
        Iterations = Iterations,
        LimitReached = false,
        Result = workingArray
      };
    }
  }

  // ===== ВСТАВКАМИ ===== (итерация = проход)
  public class InsertionSort : SortBase
  {
    public InsertionSort()
    {
      Name = "Вставками";
    }

    public override SortResult Sort(double[] input, bool ascending, Action<SortStep> onStep)
    {
      Iterations = 0;
      var workingArray = Copy(input);
      var stopwatch = Stopwatch.StartNew();
      int arrayLength = workingArray.Length;

      for (int currentIndex = 1; currentIndex < arrayLength; ++currentIndex)
      {
        ++Iterations;
        double currentValue = workingArray[currentIndex];
        int scanIndex = currentIndex - 1;

        while (scanIndex >= 0)
        {
          bool needShift = ascending
              ? workingArray[scanIndex] > currentValue
              : workingArray[scanIndex] < currentValue;
          if (!needShift) break;

          workingArray[scanIndex + 1] = workingArray[scanIndex];
          --scanIndex;
        }
        workingArray[scanIndex + 1] = currentValue;

        EmitStep(onStep, stopwatch, workingArray, scanIndex + 1, currentIndex);
      }

      stopwatch.Stop();
      return new SortResult
      {
        Name = Name,
        ElapsedMs = stopwatch.Elapsed.TotalMilliseconds,
        Iterations = Iterations,
        LimitReached = false,
        Result = workingArray
      };
    }
  }

  // ===== ШЕЙКЕРНАЯ ===== (итерация = цикл туда+обратно, макс (n-1)/2)
  public class ShakerSort : SortBase
  {
    public ShakerSort()
    {
      Name = "Шейкерная";
    }

    public override SortResult Sort(double[] input, bool ascending, Action<SortStep> onStep)
    {
      Iterations = 0;
      var workingArray = Copy(input);
      var stopwatch = Stopwatch.StartNew();
      int leftBoundary = 0;
      int rightBoundary = workingArray.Length - 1;

      while (leftBoundary < rightBoundary)
      {
        ++Iterations;
        bool swapped = false;

        for (int forwardIndex = leftBoundary; forwardIndex < rightBoundary; ++forwardIndex)
        {
          bool needSwap = ascending
              ? workingArray[forwardIndex] > workingArray[forwardIndex + 1]
              : workingArray[forwardIndex] < workingArray[forwardIndex + 1];

          if (needSwap)
          {
            (workingArray[forwardIndex], workingArray[forwardIndex + 1]) =
                (workingArray[forwardIndex + 1], workingArray[forwardIndex]);
            swapped = true;
          }
        }
        --rightBoundary;

        for (int backwardIndex = rightBoundary; backwardIndex > leftBoundary; --backwardIndex)
        {
          bool needSwap = ascending
              ? workingArray[backwardIndex - 1] > workingArray[backwardIndex]
              : workingArray[backwardIndex - 1] < workingArray[backwardIndex];

          if (needSwap)
          {
            (workingArray[backwardIndex - 1], workingArray[backwardIndex]) =
                (workingArray[backwardIndex], workingArray[backwardIndex - 1]);
            swapped = true;
          }
        }
        ++leftBoundary;

        EmitStep(onStep, stopwatch, workingArray);

        if (!swapped) break;
      }

      stopwatch.Stop();
      return new SortResult
      {
        Name = Name,
        ElapsedMs = stopwatch.Elapsed.TotalMilliseconds,
        Iterations = Iterations,
        LimitReached = false,
        Result = workingArray
      };
    }
  }

  // ===== БЫСТРАЯ ===== (итерация = обработка одного сегмента)
  public class QuickSort : SortBase
  {
    public QuickSort()
    {
      Name = "Быстрая";
    }

    private readonly Random _rng = new Random();
    private const int InsertionCutoff = 16;

    public override SortResult Sort(double[] input, bool ascending, Action<SortStep> onStep)
    {
      Iterations = 0;
      var workingArray = Copy(input);
      var stopwatch = Stopwatch.StartNew();

      var stack = new Stack<(int left, int right)>();
      stack.Push((0, workingArray.Length - 1));

      while (stack.Count > 0)
      {
        var (left, right) = stack.Pop();
        if (left >= right)
          continue;

        ++Iterations;

        if (right - left < InsertionCutoff)
        {
          for (int currentIndex = left + 1; currentIndex <= right; ++currentIndex)
          {
            double currentValue = workingArray[currentIndex];
            int scanIndex = currentIndex - 1;
            while (scanIndex >= left &&
                   (ascending ? workingArray[scanIndex] > currentValue
                              : workingArray[scanIndex] < currentValue))
            {
              workingArray[scanIndex + 1] = workingArray[scanIndex];
              --scanIndex;
            }
            workingArray[scanIndex + 1] = currentValue;
          }

          EmitStep(onStep, stopwatch, workingArray, left, right);
          continue;
        }

        int pivotIndex = Partition(workingArray, left, right, ascending, onStep, stopwatch);

        if (pivotIndex - left < right - pivotIndex)
        {
          stack.Push((pivotIndex + 1, right));
          stack.Push((left, pivotIndex - 1));
        }
        else
        {
          stack.Push((left, pivotIndex - 1));
          stack.Push((pivotIndex + 1, right));
        }
      }

      stopwatch.Stop();
      return new SortResult
      {
        Name = Name,
        ElapsedMs = stopwatch.Elapsed.TotalMilliseconds,
        Iterations = Iterations,
        LimitReached = false,
        Result = workingArray
      };
    }

    private int Partition(double[] workingArray, int leftBoundary, int rightBoundary, bool ascending,
        Action<SortStep> onStep, Stopwatch stopwatch)
    {
      int randomIndex = _rng.Next(leftBoundary, rightBoundary + 1);
      (workingArray[randomIndex], workingArray[rightBoundary]) =
          (workingArray[rightBoundary], workingArray[randomIndex]);

      double pivotValue = workingArray[rightBoundary];
      int smallerElementIndex = leftBoundary - 1;

      for (int scanIndex = leftBoundary; scanIndex < rightBoundary; ++scanIndex)
      {
        bool needSwap = ascending
            ? workingArray[scanIndex] <= pivotValue
            : workingArray[scanIndex] >= pivotValue;

        if (needSwap)
        {
          ++smallerElementIndex;
          (workingArray[smallerElementIndex], workingArray[scanIndex]) =
              (workingArray[scanIndex], workingArray[smallerElementIndex]);
        }
      }

      (workingArray[smallerElementIndex + 1], workingArray[rightBoundary]) =
          (workingArray[rightBoundary], workingArray[smallerElementIndex + 1]);

      int resultIndex = smallerElementIndex + 1;
      EmitStep(onStep, stopwatch, workingArray, resultIndex, -1, $"pivot={pivotValue:F4}");

      return resultIndex;
    }
  }

  // ===== BOGO ===== (итерация = попытка перемешивания)
  public class BogoSort : SortBase
  {
    public BogoSort()
    {
      Name = "BOGO";
    }

    public long MaxIterations { get; set; } = long.MaxValue;
    public bool LimitReached { get; private set; } = false;

    private readonly Random _rng = new Random();

    public override SortResult Sort(double[] input, bool ascending, Action<SortStep> onStep)
    {
      Iterations = 0;
      LimitReached = false;
      var workingArray = Copy(input);
      var stopwatch = Stopwatch.StartNew();

      while (!IsSorted(workingArray, ascending))
      {
        if (Iterations >= MaxIterations)
        {
          LimitReached = true;
          break;
        }
        ++Iterations;

        for (int shuffleIndex = workingArray.Length - 1; shuffleIndex > 0; --shuffleIndex)
        {
          int randomIndex = _rng.Next(shuffleIndex + 1);
          (workingArray[shuffleIndex], workingArray[randomIndex]) =
              (workingArray[randomIndex], workingArray[shuffleIndex]);
        }

        EmitStep(onStep, stopwatch, workingArray);
      }

      stopwatch.Stop();
      return new SortResult
      {
        Name = Name,
        ElapsedMs = stopwatch.Elapsed.TotalMilliseconds,
        Iterations = Iterations,
        LimitReached = LimitReached,
        Result = workingArray
      };
    }

    private bool IsSorted(double[] workingArray, bool ascending)
    {
      for (int checkIndex = 0; checkIndex < workingArray.Length - 1; ++checkIndex)
      {
        if (ascending && workingArray[checkIndex] > workingArray[checkIndex + 1]) return false;
        if (!ascending && workingArray[checkIndex] < workingArray[checkIndex + 1]) return false;
      }
      return true;
    }
  }
}
