using Sitecore.Diagnostics;
using Sitecore.Jobs;
using Sitecore.Security.Accounts;
using Sitecore.Text;
using Sitecore.WFFM.Abstractions.Dependencies;
using Sitecore.WFFM.Services.Pipelines;
using System.Text.RegularExpressions;
using System.Xml;

namespace Sitecore.Support.WFFM.Services.Pipelines.ExportToXml
{
  public class ExportFormDataToXml
  {
    public void Process(FormExportArgs args)
    {
      Job job = Context.Job;
      job?.Status.LogInfo(DependenciesManager.ResourceManager.Localize("EXPORTING_DATA"));

      // Begin of Sitecore.Support.67687
      SetDateFormat.Process(args);
      // End of Sitecore.Support.67687

      XmlDocument xmlDocument = new XmlDocument();
      xmlDocument.InnerXml = args.Packet.ToXml();
      string text = args.Parameters["contextUser"];
      Assert.IsNotNullOrEmpty(text, "contextUser");
      using (new UserSwitcher(text, true))
      {
        string exportRestriction = DependenciesManager.FormRegistryUtil.GetExportRestriction(args.Item.ID.ToString(), string.Empty);
        exportRestriction = Regex.Replace(exportRestriction, "{|}", string.Empty);
        ListString listString = new ListString(exportRestriction);
        XmlNodeList xmlNodeList = xmlDocument.SelectNodes("packet/formentry");
        Assert.IsNotNull(xmlNodeList, "roots");
        foreach (string item in listString)
        {
          foreach (XmlNode item2 in xmlNodeList)
          {
            Assert.IsNotNull(item2.Attributes, "Attributes");
            XmlAttribute xmlAttribute = item2.Attributes[item];
            if (xmlAttribute != null)
            {
              item2.Attributes.Remove(xmlAttribute);
            }
            XmlNodeList xmlNodeList2 = item2.SelectNodes($"field[@fieldid='{item.ToLower()}']");
            Assert.IsNotNull(xmlNodeList2, "nodeList");
            foreach (XmlNode item3 in xmlNodeList2)
            {
              item2.RemoveChild(item3);
            }
          }
        }
        args.Result = xmlDocument.DocumentElement.OuterXml;
      }
    }
  }
}