using Sitecore.Configuration;
using Sitecore.Data;
using Sitecore.Data.Items;
using Sitecore.WFFM.Services.Pipelines;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Sitecore.Support
{
  public static class SetDateFormat
  {
    public static void Process(FormExportArgs args)
    {
      List<Guid> dateIDs = new List<Guid>();
      Database db = Factory.GetDatabase("master");

      foreach (var id in args.Packet.Entries.SelectMany(p => p.Fields).Select(f => f.FieldId).Distinct())
      {
        Item fieldItem = db.Items[new ID(id)];
        if (fieldItem != null
            && (fieldItem["Field Link"] == "{09BF916E-79FB-4AE3-B799-659E63C75EA5}"
            || fieldItem["Field Link"] == "{95DD3FCF-2E03-4064-9968-614D1452F20B}"))
        {
          dateIDs.Add(id);
        }
      }

      var fields = args.Packet.Entries.SelectMany(p => p.Fields).Where(f => dateIDs.Contains(f.FieldId));
      foreach (var field in fields)
      {
        if (Sitecore.Form.Core.Utility.DateUtil.IsIsoDateTime(field.Value))
        {
          field.Value = Sitecore.Form.Core.Utility.DateUtil.IsoDateTimeToDateTime(field.Value).ToString(Sitecore.Configuration.Settings.GetSetting("WFM.ExportDateFormat", "dd-MMM-yyyy"));
        }
      }
    }
  }
}