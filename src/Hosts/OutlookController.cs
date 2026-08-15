using System;
using MistralOfficeAddin.Core;
using Outlook = NetOffice.OutlookApi;

namespace MistralOfficeAddin.Hosts
{
    public class OutlookController
    {
        private readonly object _rawAppObj;
        private Outlook.Application _outlookApp;

        public OutlookController(object appObj)
        {
            _rawAppObj = appObj;
        }

        private Outlook.Application GetApp()
        {
            if (_outlookApp != null) return _outlookApp;
            if (_rawAppObj == null) return null;
            try
            {
                _outlookApp = (_rawAppObj is Outlook.Application)
                    ? (Outlook.Application)_rawAppObj
                    : new Outlook.Application(null, _rawAppObj);
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("OutlookController.GetApp failed: {0}", ex.Message));
            }
            return _outlookApp;
        }

        public string GetEmailBody()
        {
            try
            {
                var app = GetApp();
                if (app != null)
                {
                    var inspector = app.ActiveInspector();
                    if (inspector != null && inspector.CurrentItem is Outlook.MailItem)
                        return ((Outlook.MailItem)inspector.CurrentItem).Body ?? string.Empty;

                    var explorer = app.ActiveExplorer();
                    if (explorer != null && explorer.Selection != null && explorer.Selection.Count > 0)
                    {
                        object item = explorer.Selection[1];
                        if (item is Outlook.MailItem)
                            return ((Outlook.MailItem)item).Body ?? string.Empty;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("OutlookController.GetEmailBody failed: {0}", ex.Message));
            }
            return string.Empty;
        }

        public string GetEmailSubject()
        {
            try
            {
                var app = GetApp();
                if (app != null)
                {
                    var inspector = app.ActiveInspector();
                    if (inspector != null && inspector.CurrentItem is Outlook.MailItem)
                        return ((Outlook.MailItem)inspector.CurrentItem).Subject ?? string.Empty;

                    var explorer = app.ActiveExplorer();
                    if (explorer != null && explorer.Selection != null && explorer.Selection.Count > 0)
                    {
                        object item = explorer.Selection[1];
                        if (item is Outlook.MailItem)
                            return ((Outlook.MailItem)item).Subject ?? string.Empty;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("OutlookController.GetEmailSubject failed: {0}", ex.Message));
            }
            return string.Empty;
        }

        public void SetComposeBody(string body)
        {
            if (body == null) body = string.Empty;
            try
            {
                var app = GetApp();
                if (app != null)
                {
                    var inspector = app.ActiveInspector();
                    if (inspector != null && inspector.CurrentItem is Outlook.MailItem)
                    {
                        ((Outlook.MailItem)inspector.CurrentItem).Body = body;
                        return;
                    }

                    var explorer = app.ActiveExplorer();
                    if (explorer != null && explorer.Selection != null && explorer.Selection.Count > 0)
                    {
                        object item = explorer.Selection[1];
                        if (item is Outlook.MailItem)
                        {
                            var reply = ((Outlook.MailItem)item).Reply();
                            reply.Body = body + "\n\n" + reply.Body;
                            reply.Display();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("OutlookController.SetComposeBody failed", ex);
                throw;
            }
        }

        public string GetActiveItemTitle()
        {
            string subject = GetEmailSubject();
            if (!string.IsNullOrEmpty(subject))
            {
                return string.Format("Email: {0}", subject);
            }
            return "OutlookSession";
        }
    }
}
