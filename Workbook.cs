using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO.Compression;
using System.Linq;
using System.Xml;

namespace Arby.XlsxReader
{
  /// <summary>
  /// Represents a (readonly) Workbook in an Excel (Open Document) file (*.xslx)
  /// </summary>
  public class Workbook : IEnumerable<Worksheet>
  {
    private readonly List<string> sharedStrings = new List<string>();
    private readonly List<Worksheet> worksheets = new List<Worksheet>();

    /// <summary>
    /// Creates a new Workbook object from a *.xlsx file
    /// </summary>
    /// <param name="filepath">The full file path and name of the xlsx file you want to read.</param>
    /// <remarks>
    /// XLSX files actually are a zipped file container with several xml (and other) files orgnanized in folders.
    /// Some (very) old xlsx documents from the first Office version(s) supporting this file format may have
    /// a deviant specification and thus cannot be read with this library. At least for now.
    /// </remarks>
    public Workbook(string filepath)
    {
      using (ZipArchive zip = ZipFile.OpenRead(filepath))
      {
        // read shared strings
        XmlNodeList texts = SelectNodesFromZipEntry(zip, "xl/sharedStrings.xml", ".//si//t");
        foreach (XmlNode nd in texts)
        {
          sharedStrings.Add(nd.InnerText);
        }

        // read Workbook & sheets
        XmlNodeList xSheets = SelectNodesFromZipEntry(zip, "xl/workbook.xml", ".//sheets//sheet");
        int sheetNo = 0;
        foreach (XmlNode nd in xSheets)
        {
          Worksheet worksheet = new Worksheet(nd.AttributeAsString("name"), nd.AttributeAsInt("sheetId"));
          worksheets.Add(worksheet);

          XmlNodeList xData = SelectNodesFromZipEntry(zip, $"xl/worksheets/sheet{++sheetNo}.xml", ".//sheetData//row");
          foreach (XmlElement xRow in xData)
          {
            int rowIndex = xRow.AttributeAsInt("r");
            string spans = xRow.AttributeAsString("spans");
            string[] spanSplit = spans.Split(':');
            if (spanSplit.Length == 2)
            {
              if (int.TryParse(spanSplit[0], out int first) && int.TryParse(spanSplit[1], out int last))
              {
                Row row = worksheet.AddRow(rowIndex, first, last);
                foreach (XmlElement xCell in xRow.GetElementsByTagName("c"))
                {
                  string cellpos = xCell.AttributeAsString("r");
                  string type = xCell.AttributeAsString("t");
                  int colIndex = Tools.GetCellPos(cellpos).ColNumber;
                  string formula = null;
                  if (xCell.GetElementsByTagName("f") is XmlNodeList flst && flst.Count > 0 && flst[0] is XmlElement xFormula)
                  {
                    formula = xFormula.InnerText;
                  }
                  object value = null;
                  string sValue = "";
                  if (xCell.GetElementsByTagName("v") is XmlNodeList vlst && vlst.Count > 0 && vlst[0] is XmlNode xValue) sValue = xValue.InnerText;
                  switch (type)
                  {
                    case "str": // string
                      value = sValue;
                      break;
                    case "s": // sharedString
                      if (int.TryParse(sValue, out int stringIndex) && sharedStrings.Count > stringIndex && stringIndex >= 0)
                      {
                        value = sharedStrings[stringIndex];
                      }
                      break;
                    case "inlineStr": // formatted (inline) string
                      if (xCell.SelectSingleNode("is") is XmlElement xIs)
                      {
                        XmlNodeList tNodes = xIs.GetElementsByTagName("t"); // we filter only the text itself and strip the formatting
                        string text = "";
                        foreach (XmlElement tNode in tNodes)
                        {
                          text += tNode.InnerText;
                        }
                        value = text;
                      }
                      break;
                    case "b": // boolean (0/1)
                      if (int.TryParse(sValue, out int iValue))
                      {
                        value = (iValue != 0);
                      }
                      else
                      {
                        value = false;
                      }
                      break;
                    default:
                      if (double.TryParse(sValue, out double dValue))
                      {
                        value = dValue;
                      }
                      break;
                  }
                  if (!string.IsNullOrEmpty(formula) || value != null)
                  {
                    row.AddCell(rowIndex, colIndex, formula, value);
                  }
                }
              }
            }
          }
        }
      }
    }

    /// <summary>
    /// Provides the Workbook as a <see cref="DataSet"/>.
    /// </summary>
    /// <returns>Returns a <see cref="DataSet"/> that represents all Worksheets in the Workbook.</returns>
    public DataSet AsDataSet()
    {
      DataSet result = new DataSet();
      foreach (Worksheet wb in this)
      {
        result.Tables.Add(wb.AsDataTable());
      }
      return result;
    }

    /// <summary>
    /// Accesses a <see cref="Worksheet"/> by its name.
    /// </summary>
    /// <param name="name">The name of the Worksheet.</param>
    /// <returns>The <see cref="Worksheet"/> with the given name or <see langword="null"/> if a worksheet with the name is not found.</returns>
    public Worksheet this[string name] => worksheets.FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.InvariantCultureIgnoreCase));

    /// <summary>
    /// Accesses a <see cref="Worksheet"/> by its index (0-based).
    /// </summary>
    /// <param name="index">Zero based index of the Worksheet.</param>
    /// <returns>The <see cref="Worksheet"/> at the given index.</returns>
    /// <exception cref="ArgumentOutOfRangeException">The index is negative or greater than <see cref="WorksheetCount"/>-1.</exception>
    public Worksheet this[int index] => worksheets.ElementAt(index);

    /// <summary>
    /// Number of <see cref="Worksheet"/>s in this Workbook.
    /// </summary>
    public int WorksheetCount => worksheets.Count;

    /// <summary>
    /// Returns an iterator that iterates through the <see cref="Worksheet"/>s in this Workbook.
    /// </summary>
    /// <returns>The iterator for the <see cref="Worksheet"/>s.</returns>
    public IEnumerator<Worksheet> GetEnumerator() => worksheets.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    // provides direct and abbreviated access to the Nodes (with a given simplified XPath) 
    // of a referenced XML file in the XLSX zip archive.
    
    /// <summary>
    /// Provides direct access to the XML nodes of a file in the XLSX container.
    /// </summary>
    /// <param name="zip">The XLSX container as already opened <see cref="ZipArchive"/>. ReadOnly access is sufficient.</param>
    /// <param name="filepath">The (relative) path and name of the XML file in the ZIP archive you want to access.</param>
    /// <param name="xpath">The XPath to the XML nodes you want to access. You don't have to bother with the default namespace as this method does it for you.</param>
    /// <returns></returns>
    private XmlNodeList SelectNodesFromZipEntry(ZipArchive zip, string filepath, string xpath)
    {
      ZipArchiveEntry zipEntry = zip.GetEntry(filepath);
      using (System.IO.Stream stream = zipEntry.Open())
      {
        XmlDocument xml = new XmlDocument();
        xml.Load(stream);
        XmlNamespaceManager mgr = new XmlNamespaceManager(xml.NameTable);
        mgr.AddNamespace("ns", xml.DocumentElement.GetNamespaceOfPrefix(""));
        xpath = xpath.Replace("//", "/");
        xpath = xpath.Replace("/", "//ns:");
        return xml.SelectNodes(xpath, mgr);
      }
    }
  }
}
