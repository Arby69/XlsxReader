using System.Globalization;

namespace System.Xml
{
  internal static class XmlExtensions
  {
    public static string AttributeAsString(this XmlNode node, string attName)
    {
      if (node.Attributes[attName] is XmlAttribute att)
      {
        return att.Value;
      }
      return "";
    }

    public static int AttributeAsInt(this XmlNode node, string attName)
    {
      int result = 0;
      if (node.Attributes[attName] is XmlAttribute att)
      {
        int.TryParse(att.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out result);
      }
      return result;
    }
  }
}
