using Sitecore;
using Sitecore.Diagnostics;
using Sitecore.Jobs;
using Sitecore.Security.Accounts;
using Sitecore.WFFM.Abstractions.Analytics;
using Sitecore.WFFM.Abstractions.Data;
using Sitecore.WFFM.Abstractions.Dependencies;
using Sitecore.WFFM.Services.Pipelines;
using Sitecore.WFFM.Speak.ViewModel;
using System;
using System.Linq;
using System.Xml;

namespace Sitecore.WFFM.Services.Pipelines.ExportToExcel
{
  public class ExportFormDataToExcel
  {
    public void Process(FormExportArgs args)
    {
      Job job = Context.Job;
      job?.Status.LogInfo(DependenciesManager.ResourceManager.Localize("EXPORTING_DATA"));
      string text = args.Parameters["contextUser"];
      Assert.IsNotNullOrEmpty(text, "contextUser");
      using (new UserSwitcher(text, true))
      {
        XmlDocument xmlDocument = new XmlDocument();
        XmlElement xmlElement = xmlDocument.CreateElement("ss:Workbook");
        XmlAttribute xmlAttribute = xmlDocument.CreateAttribute("xmlns");
        xmlAttribute.Value = "urn:schemas-microsoft-com:office:spreadsheet";
        xmlElement.Attributes.Append(xmlAttribute);
        XmlAttribute xmlAttribute2 = xmlDocument.CreateAttribute("xmlns:o");
        xmlAttribute2.Value = "urn:schemas-microsoft-com:office:office";
        xmlElement.Attributes.Append(xmlAttribute2);
        XmlAttribute xmlAttribute3 = xmlDocument.CreateAttribute("xmlns:x");
        xmlAttribute3.Value = "urn:schemas-microsoft-com:office:excel";
        xmlElement.Attributes.Append(xmlAttribute3);
        XmlAttribute xmlAttribute4 = xmlDocument.CreateAttribute("xmlns:ss");
        xmlAttribute4.Value = "urn:schemas-microsoft-com:office:spreadsheet";
        xmlElement.Attributes.Append(xmlAttribute4);
        XmlAttribute xmlAttribute5 = xmlDocument.CreateAttribute("xmlns:html");
        xmlAttribute5.Value = "http://www.w3.org/TR/REC-html40";
        xmlElement.Attributes.Append(xmlAttribute5);
        xmlDocument.AppendChild(xmlElement);
        XmlElement xmlElement2 = xmlDocument.CreateElement("Styles");
        xmlElement.AppendChild(xmlElement2);
        XmlElement xmlElement3 = xmlDocument.CreateElement("Style");
        XmlAttribute xmlAttribute6 = xmlDocument.CreateAttribute("ss", "ID", "xmlns");
        xmlAttribute6.Value = "xBoldVerdana";
        xmlElement3.Attributes.Append(xmlAttribute6);
        xmlElement2.AppendChild(xmlElement3);
        XmlElement xmlElement4 = xmlDocument.CreateElement("Font");
        XmlAttribute xmlAttribute7 = xmlDocument.CreateAttribute("ss", "Bold", "xmlns");
        xmlAttribute7.Value = "1";
        xmlElement4.Attributes.Append(xmlAttribute7);
        XmlAttribute xmlAttribute8 = xmlDocument.CreateAttribute("ss", "FontName", "xmlns");
        xmlAttribute8.Value = "verdana";
        xmlElement4.Attributes.Append(xmlAttribute8);
        xmlElement3.AppendChild(xmlElement4);
        xmlElement3 = xmlDocument.CreateElement("Style");
        xmlAttribute6 = xmlDocument.CreateAttribute("ss", "ID", "xmlns");
        xmlAttribute6.Value = "xVerdana";
        xmlElement3.Attributes.Append(xmlAttribute6);
        xmlElement2.AppendChild(xmlElement3);
        xmlElement4 = xmlDocument.CreateElement("Font");
        xmlAttribute8 = xmlDocument.CreateAttribute("ss", "FontName", "xmlns");
        xmlAttribute8.Value = "verdana";
        xmlElement4.Attributes.Append(xmlAttribute8);
        xmlElement3.AppendChild(xmlElement4);
        XmlElement xmlElement5 = xmlDocument.CreateElement("Worksheet");
        XmlAttribute xmlAttribute9 = xmlDocument.CreateAttribute("ss", "Name", "xmlns");
        xmlAttribute9.Value = "Sheet1";
        xmlElement5.Attributes.Append(xmlAttribute9);
        xmlElement.AppendChild(xmlElement5);
        XmlElement xmlElement6 = xmlDocument.CreateElement("Table");
        XmlAttribute xmlAttribute10 = xmlDocument.CreateAttribute("ss", "DefaultColumnWidth", "xmlns");
        xmlAttribute10.Value = "130";
        xmlElement6.Attributes.Append(xmlAttribute10);
        xmlElement5.AppendChild(xmlElement6);
        BuildHeader(xmlDocument, args.Item, xmlElement6);
        BuildBody(xmlDocument, args.Item, args.Packet, xmlElement6);
        XmlElement xmlElement7 = xmlDocument.CreateElement("WorksheetOptions");
        XmlElement newChild = xmlDocument.CreateElement("Selected");
        XmlElement xmlElement8 = xmlDocument.CreateElement("Panes");
        XmlElement xmlElement9 = xmlDocument.CreateElement("Pane");
        XmlElement xmlElement10 = xmlDocument.CreateElement("Number");
        xmlElement10.InnerText = "1";
        XmlElement xmlElement11 = xmlDocument.CreateElement("ActiveCol");
        xmlElement11.InnerText = "1";
        xmlElement9.AppendChild(xmlElement11);
        xmlElement9.AppendChild(xmlElement10);
        xmlElement8.AppendChild(xmlElement9);
        xmlElement7.AppendChild(xmlElement8);
        xmlElement7.AppendChild(newChild);
        xmlElement5.AppendChild(xmlElement7);
        args.Result = "<?xml version=\"1.0\"?>" + xmlDocument.InnerXml.Replace("xmlns:ss=\"xmlns\"", "");
      }
    }

    private void BuildHeader(XmlDocument doc, IFormItem item, XmlElement root)
    {
      XmlElement xmlElement = doc.CreateElement("Row");
      string exportRestriction = DependenciesManager.FormRegistryUtil.GetExportRestriction(item.ID.ToString(), string.Empty);
      if (exportRestriction.IndexOf("created", StringComparison.Ordinal) == -1)
      {
        XmlElement newChild = CreateHeaderCell("String", "Created", doc);
        xmlElement.AppendChild(newChild);
      }
      IFieldItem[] fields = item.Fields;
      foreach (IFieldItem fieldItem in fields)
      {
        if (exportRestriction.IndexOf(fieldItem.ID.ToString(), StringComparison.Ordinal) == -1)
        {
          XmlElement newChild2 = CreateHeaderCell("String", fieldItem.FieldDisplayName, doc);
          xmlElement.AppendChild(newChild2);
        }
      }
      root.AppendChild(xmlElement);
    }

    private XmlElement CreateHeaderCell(string sType, string sValue, XmlDocument doc)
    {
      XmlElement xmlElement = doc.CreateElement("Cell");
      XmlAttribute xmlAttribute = doc.CreateAttribute("ss", "StyleID", "xmlns");
      xmlAttribute.Value = "xBoldVerdana";
      xmlElement.Attributes.Append(xmlAttribute);
      XmlElement xmlElement2 = doc.CreateElement("Data");
      XmlAttribute xmlAttribute2 = doc.CreateAttribute("ss", "Type", "xmlns");
      xmlAttribute2.Value = sType;
      xmlElement2.Attributes.Append(xmlAttribute2);
      xmlElement2.InnerText = sValue;
      xmlElement.AppendChild(xmlElement2);
      return xmlElement;
    }

    private void BuildBody(XmlDocument doc, IFormItem item, FormPacket packet, XmlElement root)
    {
      foreach (FormData entry in packet.Entries)
      {
        root.AppendChild(BuildRow(entry, item, doc));
      }
    }

    private XmlElement BuildRow(FormData entry, IFormItem item, XmlDocument xd)
    {
      XmlElement xmlElement = xd.CreateElement("Row");
      string exportRestriction = DependenciesManager.FormRegistryUtil.GetExportRestriction(item.ID.ToString(), string.Empty);
      if (exportRestriction.IndexOf("created") == -1)
      {
        XmlElement newChild = CreateCell("String", entry.Timestamp.ToLocalTime().ToString("G"), xd);
        xmlElement.AppendChild(newChild);
      }
      IFieldItem[] fields = item.Fields;
      foreach (IFieldItem field in fields)
      {
        if (exportRestriction.IndexOf(field.ID.ToString(), StringComparison.Ordinal) == -1)
        {
          FieldData fieldData = entry.Fields.FirstOrDefault((FieldData f) => f.FieldId == field.ID.Guid);
          XmlElement newChild2 = CreateCell("String", (fieldData != null) ? fieldData.Value : string.Empty, xd);
          xmlElement.AppendChild(newChild2);
        }
      }
      return xmlElement;
    }

    private XmlElement CreateCell(string sType, string sValue, XmlDocument doc)
    {
      XmlElement xmlElement = doc.CreateElement("Cell");
      XmlAttribute xmlAttribute = doc.CreateAttribute("ss", "StyleID", "xmlns");
      xmlAttribute.Value = "xVerdana";
      xmlElement.Attributes.Append(xmlAttribute);
      XmlElement xmlElement2 = doc.CreateElement("Data");
      XmlAttribute xmlAttribute2 = doc.CreateAttribute("ss", "Type", "xmlns");
      xmlAttribute2.Value = sType;
      xmlElement2.Attributes.Append(xmlAttribute2);
      xmlElement2.InnerText = sValue;
      xmlElement.AppendChild(xmlElement2);
      return xmlElement;
    }
  }
}