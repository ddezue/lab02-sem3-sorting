using System;
using System.Diagnostics;

namespace SortingLab
{
  public class SortStep
  {
    public double[] Array { get; set; }
    public int Index1 { get; set; } = -1;
    public int Index2 { get; set; } = -1;
    public string Comment { get; set; } = "";
  }

  public class SortResult
  {
    public string Name { get; set; }
    public double ElapsedMs { get; set; }
    public long Iterations { get; set; }
    public bool LimitReached { get; set; }
    public double[] Result { get; set; }
  }

  public abstract class SortBase
  {
    public string Name { get; protected set; }
    public long Iterations { get; protected set; }
    public long MaxIterations { get; set; } = long.MaxValue;
    public bool LimitReached { get; protected set; } = false;

    protected bool CheckLimit()
    {
      Iterations++;
      if (Iterations >= MaxIterations)
      {
        LimitReached = true;
        return true;
      }
      return false;
    }

    public abstract SortResult Sort(double[] input, bool ascending, Action<SortStep> onStep = null);

    protected static double[] Copy(double[] source) => (double[])source.Clone();
  }

  // ---------- Пузырьковая (не более n-1 проходов) ----------
  public class BubbleSort : SortBase
  {
    public BubbleSort() { Name = "Пузырьковая"; }

    public override SortResult Sort(double[] input, bool ascending, Action<SortStep> onStep)
    {
      Iterations = 0; LimitReached = false;
      var workingArray = Copy(input);
      var stopwatch = Stopwatch.StartNew();
      int arrayLength = workingArray.Length;

      for (int outerPass = 0; outerPass < arrayLength - 1; outerPass++)
      {
        bool swapped = false;
        for (int innerIndex = 0; innerIndex < arrayLength - 1 - outerPass; innerIndex++)
        {
          if (CheckLimit())
          {
            stopwatch.Stop();
            return new SortResult { Name = Name, ElapsedMs = stopwatch.Elapsed.TotalMilliseconds, Iterations = Iterations, LimitReached = true, Result = workingArray };
          }

          bool needSwap = ascending
              ? workingArray[innerIndex] > workingArray[innerIndex + 1]
              : workingArray[innerIndex] < workingArray[innerIndex + 1];

          if (needSwap)
          {
            (workingArray[innerIndex], workingArray[innerIndex + 1]) =
                (workingArray[innerIndex + 1], workingArray[innerIndex]);
            swapped = true;
          }

          onStep?.Invoke(new SortStep
          {
            Array = Copy(workingArray),
            Index1 = innerIndex,
            Index2 = innerIndex + 1
          });
        }
        if (!swapped) break;
      }

      stopwatch.Stop();
      return new SortResult { Name = Name, ElapsedMs = stopwatch.Elapsed.TotalMilliseconds, Iterations = Iterations, LimitReached = false, Result = workingArray };
    }
  }

  // ---------- Вставками ----------
  public class InsertionSort : SortBase
  {
    public InsertionSort() { Name = "Вставками"; }

    public override SortResult Sort(double[] input, bool ascending, Action<SortStep> onStep)
    {
      Iterations = 0; LimitReached = false;
      var workingArray = Copy(input);
      var stopwatch = Stopwatch.StartNew();
      int arrayLength = workingArray.Length;

      for (int currentIndex = 1; currentIndex < arrayLength; currentIndex++)
      {
        double currentValue = workingArray[currentIndex];
        int scanIndex = currentIndex - 1;

        while (scanIndex >= 0)
        {
          if (CheckLimit())
          {
            stopwatch.Stop();
            return new SortResult { Name = Name, ElapsedMs = stopwatch.Elapsed.TotalMilliseconds, Iterations = Iterations, LimitReached = true, Result = workingArray };
          }

          bool needShift = ascending
              ? workingArray[scanIndex] > currentValue
              : workingArray[scanIndex] < currentValue;

          if (!needShift) break;

          workingArray[scanIndex + 1] = workingArray[scanIndex];
          scanIndex--;
          onStep?.Invoke(new SortStep
          {
            Array = Copy(workingArray),
            Index1 = scanIndex + 1,
            Index2 = currentIndex
          });
        }
        workingArray[scanIndex + 1] = currentValue;
        onStep?.Invoke(new SortStep
        {
          Array = Copy(workingArray),
          Index1 = scanIndex + 1,
          Index2 = currentIndex
        });
      }

      stopwatch.Stop();
      return new SortResult { Name = Name, ElapsedMs = stopwatch.Elapsed.TotalMilliseconds, Iterations = Iterations, LimitReached = false, Result = workingArray };
    }
  }

  // ---------- Шейкерная (не более (n-1)/2 проходов) ----------
  public class ShakerSort : SortBase
  {
    public ShakerSort() { Name = "Шейкерная"; }

    public override SortResult Sort(double[] input, bool ascending, Action<SortStep> onStep)
    {
      Iterations = 0; LimitReached = false;
      var workingArray = Copy(input);
      var stopwatch = Stopwatch.StartNew();
      int leftBoundary = 0;
      int rightBoundary = workingArray.Length - 1;
      int maxPasses = (workingArray.Length - 1) / 2;
      int passCounter = 0;
      bool swapped = true;

      while (swapped && leftBoundary < rightBoundary && passCounter < maxPasses)
      {
        swapped = false;
        passCounter++;

        for (int forwardIndex = leftBoundary; forwardIndex < rightBoundary; forwardIndex++)
        {
          if (CheckLimit())
          {
            stopwatch.Stop();
            return new SortResult { Name = Name, ElapsedMs = stopwatch.Elapsed.TotalMilliseconds, Iterations = Iterations, LimitReached = true, Result = workingArray };
          }

          bool needSwap = ascending
              ? workingArray[forwardIndex] > workingArray[forwardIndex + 1]
              : workingArray[forwardIndex] < workingArray[forwardIndex + 1];

          if (needSwap)
          {
            (workingArray[forwardIndex], workingArray[forwardIndex + 1]) =
                (workingArray[forwardIndex + 1], workingArray[forwardIndex]);
            swapped = true;
          }

          onStep?.Invoke(new SortStep
          {
            Array = Copy(workingArray),
            Index1 = forwardIndex,
            Index2 = forwardIndex + 1
          });
        }
        rightBoundary--;

        if (!swapped) break;
        swapped = false;

        for (int backwardIndex = rightBoundary; backwardIndex > leftBoundary; backwardIndex--)
        {
          if (CheckLimit())
          {
            stopwatch.Stop();
            return new SortResult { Name = Name, ElapsedMs = stopwatch.Elapsed.TotalMilliseconds, Iterations = Iterations, LimitReached = true, Result = workingArray };
          }

          bool needSwap = ascending
              ? workingArray[backwardIndex - 1] > workingArray[backwardIndex]
              : workingArray[backwardIndex - 1] < workingArray[backwardIndex];

          if (needSwap)
          {
            (workingArray[backwardIndex], workingArray[backwardIndex - 1]) =
                (workingArray[backwardIndex - 1], workingArray[backwardIndex]);
            swapped = true;
          }

          onStep?.Invoke(new SortStep
          {
            Array = Copy(workingArray),
            Index1 = backwardIndex - 1,
            Index2 = backwardIndex
          });
        }
        leftBoundary++;
      }

      stopwatch.Stop();
      return new SortResult { Name = Name, ElapsedMs = stopwatch.Elapsed.TotalMilliseconds, Iterations = Iterations, LimitReached = false, Result = workingArray };
    }
  }

  // ---------- Быстрая ----------
  public class QuickSort : SortBase
  {
    public QuickSort() { Name = "Быстрая"; }

    public override SortResult Sort(double[] input, bool ascending, Action<SortStep> onStep)
    {
      Iterations = 0; LimitReached = false;
      var workingArray = Copy(input);
      var stopwatch = Stopwatch.StartNew();

      QuickSortRecursive(workingArray, 0, workingArray.Length - 1, ascending, onStep);

      stopwatch.Stop();
      return new SortResult { Name = Name, ElapsedMs = stopwatch.Elapsed.TotalMilliseconds, Iterations = Iterations, LimitReached = LimitReached, Result = workingArray };
    }

    private void QuickSortRecursive(double[] workingArray, int leftBoundary, int rightBoundary, bool ascending, Action<SortStep> onStep)
    {
      if (leftBoundary < rightBoundary && !LimitReached)
      {
        if (CheckLimit())
        {
          LimitReached = true;
          return;
        }

        int pivotIndex = Partition(workingArray, leftBoundary, rightBoundary, ascending, onStep);
        if (LimitReached) return;

        QuickSortRecursive(workingArray, leftBoundary, pivotIndex - 1, ascending, onStep);
        QuickSortRecursive(workingArray, pivotIndex + 1, rightBoundary, ascending, onStep);
      }
    }

    private int Partition(double[] workingArray, int leftBoundary, int rightBoundary, bool ascending, Action<SortStep> onStep)
    {
      double pivotValue = workingArray[rightBoundary];
      int smallerElementIndex = leftBoundary - 1;

      for (int scanIndex = leftBoundary; scanIndex < rightBoundary; scanIndex++)
      {
        if (LimitReached) return smallerElementIndex + 1;

        bool needSwap = ascending
            ? workingArray[scanIndex] <= pivotValue
            : workingArray[scanIndex] >= pivotValue;

        if (needSwap)
        {
          smallerElementIndex++;
          (workingArray[smallerElementIndex], workingArray[scanIndex]) =
              (workingArray[scanIndex], workingArray[smallerElementIndex]);
        }

        onStep?.Invoke(new SortStep
        {
          Array = Copy(workingArray),
          Index1 = smallerElementIndex,
          Index2 = scanIndex,
          Comment = $"pivot={pivotValue:F4}"
        });
      }

      (workingArray[smallerElementIndex + 1], workingArray[rightBoundary]) =
          (workingArray[rightBoundary], workingArray[smallerElementIndex + 1]);

      onStep?.Invoke(new SortStep
      {
        Array = Copy(workingArray),
        Index1 = smallerElementIndex + 1,
        Index2 = rightBoundary,
        Comment = $"pivot={pivotValue:F4}"
      });

      return smallerElementIndex + 1;
    }
  }

  // ---------- BOGO ----------
  public class BogoSort : SortBase
  {
    public BogoSort() { Name = "BOGO"; }

    public override SortResult Sort(double[] input, bool ascending, Action<SortStep> onStep)
    {
      Iterations = 0; LimitReached = false;
      var workingArray = Copy(input);
      var stopwatch = Stopwatch.StartNew();
      var randomGenerator = new Random();

      while (!IsSorted(workingArray, ascending))
      {
        if (CheckLimit())
        {
          stopwatch.Stop();
          return new SortResult { Name = Name, ElapsedMs = stopwatch.Elapsed.TotalMilliseconds, Iterations = Iterations, LimitReached = true, Result = workingArray };
        }

        for (int shuffleIndex = workingArray.Length - 1; shuffleIndex > 0; shuffleIndex--)
        {
          int randomIndex = randomGenerator.Next(shuffleIndex + 1);
          (workingArray[shuffleIndex], workingArray[randomIndex]) =
              (workingArray[randomIndex], workingArray[shuffleIndex]);
        }

        onStep?.Invoke(new SortStep { Array = Copy(workingArray) });
      }

      stopwatch.Stop();
      return new SortResult { Name = Name, ElapsedMs = stopwatch.Elapsed.TotalMilliseconds, Iterations = Iterations, LimitReached = false, Result = workingArray };
    }

    private bool IsSorted(double[] workingArray, bool ascending)
    {
      for (int checkIndex = 0; checkIndex < workingArray.Length - 1; checkIndex++)
      {
        if (ascending && workingArray[checkIndex] > workingArray[checkIndex + 1]) return false;
        if (!ascending && workingArray[checkIndex] < workingArray[checkIndex + 1]) return false;
      }
      return true;
    }
  }
}