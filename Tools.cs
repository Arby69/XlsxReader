namespace Arby.XlsxReader
{
  internal static class Tools
  {
    internal static CellPos GetCellPos(string cellpos)
    {
      int colIndex = 0;
      int p = 0;
      foreach (char c in cellpos)
      {
        if (c >= 65 && c <= 90)
        {
          colIndex *= 26;
          colIndex += c - 64;
        }
        else if (c >= 97 && c <= 122)
        {
          colIndex *= 26;
          colIndex += c - 96;
        }
        else break;
        p++;
      }
      int.TryParse(cellpos.Substring(p), out int rowIndex);
      return new CellPos(rowIndex, colIndex);
    }

    internal static string GetColumnString(int column)
    {
      string colName = "";
      do
      {
        int n = column % 26;
        column /= 26;
        colName = (char)(n + 64) + colName;
      }
      while (column != 0);
      return colName;
    }

    internal static string GetPosString(int row, int column) => $"{GetColumnString(column)}{row}";
  }
}
