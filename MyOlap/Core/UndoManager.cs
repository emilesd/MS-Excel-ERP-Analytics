namespace MyOlap.Core;

/// <summary>
/// Maintains up to 3 previous view states so the user can revert
/// (Product Brief: "Undo Last – revert back to a previous view upto 3 times").
/// </summary>
public class UndoManager
{
    private const int MaxUndoLevels = 3;
    private readonly LinkedList<ViewState> _stack = new();

    public void Push(ViewState state) 
    {
        _stack.AddLast(state.Clone());
        while (_stack.Count > MaxUndoLevels)
            _stack.RemoveFirst();
    }

    public ViewState? Pop()
    {
        if (_stack.Count == 0) return null;
        var last = _stack.Last!.Value;
        _stack.RemoveLast();
        return last;
    }

    public bool CanUndo => _stack.Count > 0;
    public int Count => _stack.Count;

    public void Clear() => _stack.Clear();
}
