using System.Collections;
using System.Collections.Generic;

namespace Arby.XlsxReader
{
  /// <summary>
  /// A row of cells on a <see cref="Worksheet"/>
  /// </summary>
  public class Row : IEnumerable<Cell>
  {
    private readonly Dictionary<int, Cell> cells = new Dictionary<int, Cell>();

    internal Row(int rowIndex, int firstColumn, int lastColumn)
    {
      RowNumber = rowIndex;
      FirstColumn = firstColumn;
      LastColumn = lastColumn;
    }

    internal Cell AddCell(int row, int column, string formula, object value)
    {
      Cell cell = new Cell(row, column, formula, value);
      cells[column] = cell;
      return cell;
    }

    /// <summary>
    /// The first column that reportedly has data (but mostly we have more columns than needed)
    /// </summary>
    public int FirstColumn { get; }

    /// <summary>
    /// The last column that reportedly has data (but mostly we have more columns than needed)
    /// </summary>
    public int LastColumn { get; }

    /// <summary>
    /// Number of reported columns (result of <see cref="FirstColumn"/> and <see cref="LastColumn"/>)
    /// </summary>
    public int ColumnCount => LastColumn - FirstColumn + 1;

    /// <summary>
    /// The 1-based number of the row on the worksheet.
    /// </summary>
    public int RowNumber { get; }

    /// <summary>
    /// The number of actual cells.
    /// </summary>
    /// <remarks>
    /// This number might differ from <see cref="ColumnCount"/> (which sometimes is greater), as empty cells are not explicitly listed.
    /// </remarks>
    public int CellCount => cells.Count;

    /// <summary>
    /// Gets the cell from the column with the number <paramref name="column"/>"/>
    /// </summary>
    /// <param name="column">The 1-based number of the column</param>
    /// <returns>The <see cref="Cell"/> of the specified column or <see langword="null"/> when the cell is not specified or empty.</returns>
    /// <remarks>You are allowed to use negative column numbers or numbers that exceed the column count. In this case, <see langword="null"/> is returned, too.</remarks>
    public Cell this[int column] => GetCell(column);

    /// <summary>
    /// Gets the cell from the column with the number <paramref name="column"/>"/>
    /// </summary>
    /// <param name="column">The 1-based number of the column</param>
    /// <returns>The <see cref="Cell"/> of the specified column or <see langword="null"/> when the cell is not specified or empty.</returns>
    /// <remarks>You are allowed to use negative column numbers or numbers that exceed the column count. In this case, <see langword="null"/> is returned, too.</remarks>
    public Cell GetCell(int column)
    {
      if (cells.TryGetValue(column, out Cell cell))
      {
        return cell;
      }
      return null;
    }

    /// <summary>
    /// Returns an enumerator that iterates through the (defined, not empty) cells.
    /// </summary>
    /// <returns>A <see cref="IEnumerator{Cell}"/> for the cells in this row.</returns>
    public IEnumerator<Cell> GetEnumerator()
    {
      return cells.Values.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
  }
}
