<%@ WebHandler Language="C#" Class="Consignment.KeepAlive" %>

using System;
using System.Web;
using System.Web.SessionState;

namespace Consignment
{
    /// <summary>
    /// Lightweight ping used by the master page while a user has a form open.
    /// Every hit renews the user's session and keeps the IIS worker process from
    /// going idle, so long data entry (e.g. 10-15 invoices on one LR) does not log the user out.
    /// Returns "1" when the session is alive, "0" when it has already expired.
    /// </summary>
    public class KeepAlive : IHttpHandler, IRequiresSessionState
    {
        public void ProcessRequest(HttpContext context)
        {
            context.Response.Cache.SetCacheability(HttpCacheability.NoCache);
            context.Response.Cache.SetNoStore();
            context.Response.ContentType = "text/plain";
            context.Response.Write(context.Session != null && context.Session["UserId"] != null ? "1" : "0");
        }

        public bool IsReusable { get { return true; } }
    }
}
