namespace ExcelDataReader.Core.OpenXmlFormat;

internal sealed class OpenXmlRowOrder
{
    private List<Run>? _runs;
    private Run _current;
    private int _lastIndex;
    private int _lastOrdinal;
    private int _physicalCount;
    private bool _descending;

    public void Add(int index)
    {
        int ordinal = _physicalCount++;
        if (_current.Count == 0)
        {
            _current = new Run(index, ordinal, 0, 0, 1);
            _lastIndex = index;
            _lastOrdinal = ordinal;
            return;
        }

        if (index == _lastIndex)
            return;

        // Compress logical row groups while retaining their physical fragment ordinals.
        long step = (long)index - _lastIndex;
        int ordinalStep = ordinal - _lastOrdinal;
        _descending |= step < 0;
        if (step >= int.MinValue && step <= int.MaxValue &&
            (_current.Count == 1 || (step == _current.Step && ordinalStep == _current.OrdinalStep)))
        {
            _current = _current with { Step = (int)step, OrdinalStep = ordinalStep, Count = _current.Count + 1 };
        }
        else
        {
            (_runs ??= []).Add(_current);
            _current = new Run(index, ordinal, 0, 0, 1);
        }

        _lastIndex = index;
        _lastOrdinal = ordinal;
    }

    public Region[] Build()
    {
        if (!_descending)
            return [];

        int runCount = (_runs?.Count ?? 0) + 1;
        var suffixMinimum = new int[runCount + 1];
        int minimum = int.MaxValue;
        suffixMinimum[runCount] = minimum;
        for (int i = runCount - 1; i >= 0; i--)
        {
            var run = GetRun(i);
            minimum = Math.Min(minimum, Math.Min(run.First, run.Last));
            suffixMinimum[i] = minimum;
        }

        // A boundary is safe only when the prefix maximum is strictly below the
        // remaining suffix minimum. This keeps unseen lower rows and later equal
        // fragments inside the same region until they can all be emitted in order.
        List<Region> regions = [];
        int maximum = int.MinValue;
        int start = 0;
        int previous = 0;
        bool descending = false;
        for (int i = 0; i < runCount; i++)
        {
            var run = GetRun(i);
            for (int offset = 0; offset < run.Count; offset++)
            {
                int ordinal = run.Ordinal + offset * run.OrdinalStep;
                int index = (int)((long)run.First + (long)offset * run.Step);
                maximum = Math.Max(maximum, index);
                descending |= ordinal > start && index < previous;
                previous = index;

                bool moreInRun = offset + 1 < run.Count;
                int nextMinimum = moreInRun
                    ? Math.Min(run.Step < 0 ? run.Last : index + run.Step, suffixMinimum[i + 1])
                    : suffixMinimum[i + 1];

                // Equal indices cannot straddle a boundary: later fragments still
                // belong to the same logical row, even when other rows intervene.
                if ((moreInRun || i + 1 < runCount) && maximum >= nextMinimum)
                    continue;

                int nextOrdinal = moreInRun
                    ? ordinal + run.OrdinalStep
                    : i + 1 < runCount ? GetRun(i + 1).Ordinal : _physicalCount;
                if (descending)
                    regions.Add(new Region(start, nextOrdinal - 1));
                start = nextOrdinal;
                descending = false;
            }
        }

        return [.. regions];
    }

    private Run GetRun(int index) => _runs != null && index < _runs.Count ? _runs[index] : _current;

    internal readonly record struct Region(int Start, int End);

    // Count counts logical row groups, not fragments. Ordinal and OrdinalStep
    // locate those groups in the physical fragment stream, including repetitions.
    private readonly record struct Run(int First, int Ordinal, int Step, int OrdinalStep, int Count)
    {
        public int Last => (int)(First + (Count - 1L) * Step);
    }
}