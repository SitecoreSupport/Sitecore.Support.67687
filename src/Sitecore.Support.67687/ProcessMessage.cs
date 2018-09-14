using Sitecore.Data;
using Sitecore.Data.Items;
using Sitecore.Form.Core.Configuration;
using Sitecore.Form.Core.Controls.Data;
using Sitecore.Form.Core.Pipelines.ProcessMessage;
using Sitecore.Form.Core.Utility;
using Sitecore.Forms.Core.Data;
using Sitecore.Links;
using Sitecore.StringExtensions;
using Sitecore.Web;
using System;
using System.Net.Mail;
using System.Text.RegularExpressions;

namespace Sitecore.Form.Core.Pipelines.ProcessMessage.ProcessMessage
{
  public class ProcessMessage
  {
    private static readonly string SrcReplacer = string.Join(string.Empty, "src=\"", Sitecore.Web.WebUtil.GetServerUrl(), "/~");

    private static readonly string ShortHrefReplacer = string.Join(string.Empty, "href=\"", Sitecore.Web.WebUtil.GetServerUrl(), "/");

    private static readonly string ShortHrefMediaReplacer = string.Join(string.Empty, "href=\"", Sitecore.Web.WebUtil.GetServerUrl(), "/~/");

    private static readonly string HrefReplacer = ShortHrefReplacer + "~";

    public void ExpandLinks(ProcessMessageArgs args)
    {
      string value = LinkManager.ExpandDynamicLinks(args.Mail.ToString());
      args.Mail.Remove(0, args.Mail.Length);
      args.Mail.Append(value);
    }

    public void ExpandTokens(ProcessMessageArgs args)
    {
      foreach (AdaptedControlResult field in args.Fields)
      {
        FieldItem fieldItem = new FieldItem(StaticSettings.ContextDatabase.GetItem(field.FieldID));
        string value = field.Value;
        value = FieldReflectionUtil.GetAdaptedValue(fieldItem, value);
        value = Regex.Replace(value, "src=\"/sitecore/shell/themes/standard/~", SrcReplacer);
        value = Regex.Replace(value, "href=\"/sitecore/shell/themes/standard/~", HrefReplacer);
        value = Regex.Replace(value, "on\\w*=\".*?\"", string.Empty);
        if (args.MessageType == MessageType.SMS)
        {
          args.Mail.Replace(Sitecore.StringExtensions.StringExtensions.FormatWith("[{0}]", fieldItem.FieldDisplayName), value);
          args.Mail.Replace(Sitecore.StringExtensions.StringExtensions.FormatWith("[{0}]", fieldItem.Name), value);
        }
        else
        {
          if (!string.IsNullOrEmpty(field.Parameters) && field.Parameters.StartsWith("multipleline") && args.IsBodyHtml)
          {
            value = value.Replace(Environment.NewLine, "<br/>");
          }
          args.Mail.Replace(Sitecore.StringExtensions.StringExtensions.FormatWith("[<label id=\"{0}\">{1}</label>]", fieldItem.ID, fieldItem.Title), value);
          args.Mail.Replace(Sitecore.StringExtensions.StringExtensions.FormatWith("[<label id=\"{0}\">{1}</label>]", fieldItem.ID, fieldItem.Name), value);
        }
        args.From = args.From.Replace("[" + fieldItem.ID + "]", field.Value);
        args.To.Replace(string.Join(string.Empty, "[", fieldItem.ID.ToString(), "]"), field.Value);
        args.CC.Replace(string.Join(string.Empty, "[", fieldItem.ID.ToString(), "]"), field.Value);
        args.Subject.Replace(string.Join(string.Empty, "[", fieldItem.ID.ToString(), "]"), field.Value);
        args.From = args.From.Replace("[" + fieldItem.FieldDisplayName + "]", field.Value);
        args.To.Replace(string.Join(string.Empty, "[", fieldItem.FieldDisplayName, "]"), field.Value);
        args.CC.Replace(string.Join(string.Empty, "[", fieldItem.FieldDisplayName, "]"), field.Value);
        args.Subject.Replace(string.Join(string.Empty, "[", fieldItem.FieldDisplayName, "]"), field.Value);
        args.From = args.From.Replace("[" + field.FieldName + "]", field.Value);
        args.To.Replace(string.Join(string.Empty, "[", field.FieldName, "]"), field.Value);
        args.CC.Replace(string.Join(string.Empty, "[", field.FieldName, "]"), field.Value);
        args.Subject.Replace(string.Join(string.Empty, "[", field.FieldName, "]"), field.Value);
      }
    }

    public void AddHostToItemLink(ProcessMessageArgs args)
    {
      args.Mail.Replace("href=\"/", ShortHrefReplacer);
    }

    public void AddHostToMediaItem(ProcessMessageArgs args)
    {
      args.Mail.Replace("href=\"~/", ShortHrefMediaReplacer);
    }

    public void AddAttachments(ProcessMessageArgs args)
    {
      foreach (AdaptedControlResult field in args.Fields)
      {
        if (!string.IsNullOrEmpty(field.Parameters) && field.Parameters.StartsWith("medialink") && !string.IsNullOrEmpty(field.Value))
        {
          ItemUri itemUri = ItemUri.Parse(field.Value);
          if (itemUri != (ItemUri)null)
          {
            Item item = Database.GetItem(itemUri);
            if (item != null)
            {
              MediaItem mediaItem = new MediaItem(item);
              args.Attachments.Add(new Attachment(mediaItem.GetMediaStream(), string.Join(".", mediaItem.Name, mediaItem.Extension), mediaItem.MimeType));
            }
          }
        }
      }
    }

    public void BuildToFromRecipient(ProcessMessageArgs args)
    {
      if (!string.IsNullOrEmpty(args.Recipient) && !string.IsNullOrEmpty(args.RecipientGateway))
      {
        if (args.To.Length > 0)
        {
          args.To.Remove(0, args.To.Length);
        }
        args.To.Append(args.Fields.GetValueByFieldID(args.Recipient)).Append(args.RecipientGateway);
      }
    }

    public void SendEmail(ProcessMessageArgs args)
    {
      SmtpClient smtpClient = new SmtpClient(args.Host);
      smtpClient.EnableSsl = args.EnableSsl;
      SmtpClient smtpClient2 = smtpClient;
      if (args.Port != 0)
      {
        smtpClient2.Port = args.Port;
      }
      smtpClient2.Credentials = args.Credentials;
      smtpClient2.Send(GetMail(args));
    }

    private MailMessage GetMail(ProcessMessageArgs args)
    {
      MailMessage mail = new MailMessage(args.From, args.To.ToString(), args.Subject.ToString(), args.Mail.ToString())
      {
        IsBodyHtml = args.IsBodyHtml
      };
      if (args.CC.Length > 0)
      {
        mail.CC.Add(args.CC.ToString());
      }
      if (args.BCC.Length > 0)
      {
        mail.Bcc.Add(args.BCC.ToString());
      }
      args.Attachments.ForEach(delegate (Attachment attachment)
      {
        mail.Attachments.Add(attachment);
      });
      return mail;
    }
  }
}