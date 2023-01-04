namespace Arby.XlsxReader
{
  /// <summary>
  /// Represents a position tuple (row and column) of a <see cref="Cell"/>.
  /// </summary>
  public struct CellPos
  {
    /// <summary>
    /// The 1-based row number of a cell position
    /// </summary>
    public readonly int RowNumber;

    /// <summary>
    /// The 1-based column number of a cell position
    /// </summary>
    public readonly int ColNumber;

    internal CellPos(int row, int column)
    {
      RowNumber = row;
      ColNumber = column;
    }

    /// <summary>
    /// The string representation of a cell position that is typical for Excel.
    /// </summary>
    public string PosString => Tools.GetPosString(RowNumber, ColNumber);
  }
}
