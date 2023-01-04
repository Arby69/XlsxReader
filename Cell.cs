using System;

namespace Arby.XlsxReader
{
  /// <summary>
  /// Represents a single cell in a <see cref="Row"/> on a <see cref="Worksheet"/>.
  /// </summary>
  public class Cell
  {
    internal Cell(int row, int column, string formula, object value)
    {
      CellPos = new CellPos(row, column);
      Value = value;
      Formula = formula;
    }

    /// <summary>
    /// Gives the position of the cell on the worksheet
    /// </summary>
    public CellPos CellPos { get; }

    /// <summary>
    /// The row number of the cell on the worksheet (starts with 1)
    /// </summary>
    public int RowNumber => CellPos.RowNumber;

    /// <summary>
    /// The column number of the cell on the worksheet (starts with 1)
    /// </summary>
    public int ColNumber => CellPos.ColNumber;

    /// <summary>
    /// The (calculated) content value of the cell
    /// </summary>
    /// <remarks>
    /// The <see cref="Type"/> of the value can be <see cref="string"/>, <see cref="double"/> or <see cref="bool"/>.
    /// If the cell content represents a date, time or <see cref="DateTime"/>, it will be a numeric (<see cref="double"/>) 
    /// value you can use with the <see cref="DateTime.FromOADate"/> method to get a <see cref="DateTime"/> value.
    /// </remarks>
    public object Value { get; }

    /// <summary>
    /// The formula (if exists) that results in the value
    /// </summary>
    /// <remarks>
    /// Formula sequences are not yet fully implemented, so some cells might have no formula though they are calculated by one in Excel.
    /// </remarks>
    public string Formula { get; }

    /// <summary>
    /// Returns a string that represents the current object.
    /// </summary>
    /// <returns>A string that represents the current object.</returns>
    public override string ToString()
    {
      return $"{CellPos.PosString} = \"{Value}\" ({Value.GetType().Name})";
    }
  }
}
