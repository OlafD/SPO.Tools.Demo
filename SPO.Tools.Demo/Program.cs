using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Security;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using SPO.Tools;
using SPO.Tools.Extensions;

namespace SPO.Tools.Demo
{
    class Program
    {
        static readonly string DEMO_USER = "adelev@mod681489.onmicrosoft.com";

        static void Main(string[] args)
        {
            var siteUrl = "https://mod681489.sharepoint.com/sites/SPO.Tools-Demo";

            string user = "admin@mod681489.onmicrosoft.com";
            var pw = GetPassword();

            using (ClientContext ctx = new ClientContext(siteUrl))
            {
                ctx.Credentials = new SharePointOnlineCredentials(user, pw);

                Web web = ctx.Web;

                ctx.Load(web);
                ctx.ExecuteQuery();

                SimpleErrorlogDemo(ctx);
                MailTemplateDemo(ctx);

                Console.WriteLine("Switch context to admin site");

                var tenantAdminUrl = NewSiteUrl.GetTenantAdminSite(siteUrl);

                using (ClientContext tenantCtx = new ClientContext(tenantAdminUrl))
                {
                    tenantCtx.Credentials = new SharePointOnlineCredentials(user, pw);

                    Web rootWeb = tenantCtx.Web;

                    tenantCtx.Load(rootWeb);
                    tenantCtx.ExecuteQuery();

                    string url = NewSiteUrlDemo(tenantCtx);

                    SimpleErrorlog errorlog = new SimpleErrorlog(ctx);
                    errorlog.SetCorrelation();
                    errorlog.WriteInformation("NewSiteUrlDemo", String.Format("Generated url: {0}", url));
                }
            }
        }

        static void SimpleErrorlogDemo(ClientContext ctx)
        {
            SimpleErrorlog errorlog = new SimpleErrorlog(ctx);

            errorlog.EnsureList();

            errorlog.SetCorrelation();

            errorlog.WriteInformation("SimpleErrorlogDemo", "Demo information message");

            errorlog.WriteWarning("SimpleErrorlogDemo", "Demo warning message");

            errorlog.WriteError("SimpleErrorlogDemo", "Demo error message");

            try
            {
                var i = 10;
                var j = 0;
                var k = i / j;
            }
            catch (Exception ex)
            {
                errorlog.WriteError("SimpleErrorlogDemo", ex);
            }

            Console.WriteLine("Errorlog demo done.");
        }

        static void MailTemplateDemo(ClientContext ctx)
        {
            MailTemplate.EnsureList(ctx); // just necessary one time, should be done during site provisioning

            AddMailTemplate(ctx);

            Web web = ctx.Web;
            ctx.Load(web);
            ctx.ExecuteQueryRetry();

            User user = web.EnsureUser(DEMO_USER);
            ctx.Load(user);
            ctx.ExecuteQueryRetry();

            User manager = user.GetManager();

            MailTemplate mt = new MailTemplate(ctx, "Demo Template");
            mt.SetTokenValue("user", manager.Title);
            var subject = mt.GetMailSubject();
            var body = mt.GetMailBody();

            EmailProperties mailProps = new EmailProperties();
            mailProps.To = new string[] { manager.Email };
            mailProps.Subject = subject;
            mailProps.Body = body;
            Utility.SendEmail(ctx, mailProps);
            ctx.ExecuteQueryRetry();

            Console.WriteLine("Mail sent.");
        }

        static string NewSiteUrlDemo(ClientContext ctx)
        {
            NewSiteUrl newUrl = new NewSiteUrl(ctx);

            var url = newUrl.GetNewFullSiteUrl();

            Console.WriteLine("Generated url: {0}", url);

            return url;
        }

        static void AddMailTemplate(ClientContext ctx)
        {
            List list = ctx.Web.Lists.GetByTitle(MailTemplate.ListName);
            ctx.Load(list);
            ctx.ExecuteQueryRetry();

            ListItemCreationInformation newItem = new ListItemCreationInformation();

            ListItem item = list.AddItem(newItem);

            item["Title"] = "Demo Template";
            item["Subject"] = "Message for [*user*]";
            item["Body"] = "Hello [*user*]<br/><br/>This is a demo message.<br/><br/>Regards<br/>BAFH";
            item.Update();

            ctx.ExecuteQueryRetry();
        }

        static SecureString GetPassword()
        {
            Console.Write("Enter password: ");

            var pwd = new SecureString();
            while (true)
            {
                ConsoleKeyInfo i = Console.ReadKey(true);
                if (i.Key == ConsoleKey.Enter)
                {
                    break;
                }
                else if (i.Key == ConsoleKey.Backspace)
                {
                    if (pwd.Length > 0)
                    {
                        pwd.RemoveAt(pwd.Length - 1);
                        Console.Write("\b \b");
                    }
                }
                else
                {
                    pwd.AppendChar(i.KeyChar);
                    Console.Write("*");
                }
            }

            Console.WriteLine();

            return pwd;
        }
    }
}
