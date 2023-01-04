using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Arby.XlsxReader
{
  /// <summary>
  /// Represents a worksheet in an Excel (Open Document) file (<see cref="Workbook"/>).
  /// </summary>
  public class Worksheet : IEnumerable<Cell>
  {
    private readonly Dictionary<int, Row> rows = new Dictionary<int, Row>();

    internal Worksheet(string name, int id)
    {
      Name = name;
      Id = id;
    }

    internal Row AddRow(int index, int firstColumn, int lastColumn)
    {
      Row row = new Row(index, firstColumn, lastColumn);
      rows[index] = row;
      return row;
    }

    /// <summary>
    /// The display name from the tab of the worksheet in the <see cref="Workbook"/>.
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// The Id of the worksheet.
    /// </summary>
    public int Id { get; }

    /// <summary>
    /// An index accessor for the cells on this worksheet
    /// </summary>
    /// <param name="row">Row number (1-based)</param>
    /// <param name="column">Column number (1-based)</param>
    /// <returns>
    /// A <see cref="Cell"/> object that represents the cell in the given position on the Worksheet.
    /// Returns <see langword="null"/> when the position is outside the worksheet's range or the cell is empty.
    /// </returns>
    public Cell this[int row, int column] => GetRow(row)?.GetCell(column);

    /// <summary>
    /// Retrieve a <see cref="Row"/> of a worksheet.
    /// </summary>
    /// <param name="rowNumber">Row number (1-based)</param>
    /// <returns>The <see cref="Row"/> with the given number or <see langword="null"/> if the number is outside the worksheet's range.</returns>
    public Row GetRow(int rowNumber)
    {
      if (rows.TryGetValue(rowNumber, out Row row))
      {
        return row;
      }
      return null;
    }

    /// <summary>
    /// Retrieve the value of a <see cref="Cell"/> on the Worksheet.
    /// </summary>
    /// <param name="row">Row number (1-based)</param>
    /// <param name="column">Column number (1-based)</param>
    /// <returns>
    /// The value of the <see cref="Cell"/> in the given position on the Worksheet.
    /// The return value is <see cref="System.DBNull"/> if the cell is empty.
    /// The return value is <see langword="null"/> if the position is outside the Worksheet's range.
    /// </returns>
    /// <remarks>
    /// The <see cref="Type"/> of the value can be <see cref="string"/>, <see cref="double"/> or <see cref="bool"/>.
    /// If the cell content represents a date, time or <see cref="DateTime"/>, it will be a numeric (<see cref="double"/>) 
    /// value you can use with the <see cref="DateTime.FromOADate"/> method to get a <see cref="DateTime"/> value.
    /// </remarks>
    public object GetValue(int row, int column) => this[row, column]?.Value;

    /// <summary>
    /// Creates a <see cref="DataTable"/> of the Worksheet data in a given Range.
    /// </summary>
    /// <param name="startRow">Number of first row (see remarks)</param>
    /// <param name="endRow">Number of last row (see remarks)</param>
    /// <param name="firstColumn">Number of first column (see remarks)</param>
    /// <param name="lastColumn">Number of last column (see remarks)</param>
    /// <returns>The data of this Worksheet in that Range as a <see cref="DataTable"/>.</returns>
    /// <remarks>
    /// You may clip the worksheet data to a given range (by the parameters). This results in a smaller
    /// <see cref="DataTable"/>. You must consider that the row and column indices in that DataTable then
    /// depend on the clipped Range.
    /// </remarks>
    public DataTable AsDataTable(int startRow, int endRow, int firstColumn, int lastColumn)
    {
      DataTable result = new DataTable(Name);
      for (int i = firstColumn; i <= lastColumn; i++)
      {
        var col = result.Columns.Add(Tools.GetColumnString(i));
        col.DataType = typeof(object);
      }
      for (int i = startRow; i <= endRow; i++)
      {
        result.Rows.Add(result.NewRow());
      }
      foreach (Cell cell in this)
      {
        result.Rows[cell.RowNumber - startRow][cell.ColNumber - firstColumn] = cell.Value;
      }
      return result;
    }

    /// <summary>
    /// Creates a <see cref="DataTable"/> for easy access to the Worksheet data.
    /// </summary>
    /// <returns>The data of this Worksheet as a <see cref="DataTable"/>.</returns>
    public DataTable AsDataTable()
    {
      return AsDataTable(rows.Min(r => r.Key), rows.Max(r => r.Key), rows.Min(r => r.Value.FirstColumn), rows.Max(r => r.Value.LastColumn));
    }

    /// <summary>
    /// Creates a <see cref="DataTable"/> of the Worksheet data in a given Range.
    /// </summary>
    /// <param name="range">The Range string of the desired worksheet's Range (eg. "A1:E10").</param>
    /// <returns>The data of this Worksheet in that Range as a <see cref="DataTable"/>.</returns>
    public DataTable AsDataTable(string range)
    {
      string[] tokens = range.Split(':');
      if (tokens.Length < 1) throw new ArgumentException("Invalid Range", nameof(range));
      CellPos p1 = Tools.GetCellPos(tokens[0]);
      CellPos p2 = p1;
      if (tokens.Length > 1) p2 = Tools.GetCellPos(tokens[1]);
      return AsDataTable(Math.Min(p1.RowNumber, p2.RowNumber), Math.Max(p1.RowNumber, p2.RowNumber), Math.Min(p1.ColNumber, p2.ColNumber), Math.Max(p1.ColNumber, p2.ColNumber));
    }

    /// <summary>
    /// Returns an iterator that iterates through all defined <see cref="Cell"/>s on this Worksheet.
    /// </summary>
    /// <returns>The iterator for the Worksheet's <see cref="Cell"/>s.</returns>
    public IEnumerator<Cell> GetEnumerator()
    {
      foreach (Row row in rows.Values)
      {
        foreach (Cell cell in row)
        {
          yield return cell;
        }
      }
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
  }
}
