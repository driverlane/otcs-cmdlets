using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Net;
using System.Text;
using Newtonsoft.Json.Linq;

namespace cscmdlets
{
    internal static class RestApi
    {

        private static String ticket;
        private static int expiry = 28;
        private static DateTime expires;

        internal static Boolean CheckConnection()
        {

            // if there's no previous connection or an expired one open a new one
            if (expires == null || DateTime.Now > expires)
                OpenConnection();

            // return a response if we've got a ticket
            if (String.IsNullOrEmpty(ticket))
                return false;
            else
                return true;

        }

        internal static void OpenConnection()
        {
            // check the url
            if (!Globals.RestUrl.EndsWith("/"))
                Globals.RestUrl = String.Format("{0}/", Globals.RestUrl);

            // build the url and parameters
            string url = String.Format("{0}auth/", Globals.RestUrl);
            NameValueCollection parms = new NameValueCollection();
            parms.Add("username", Globals.Username);
            parms.Add("password", Globals.Password);
            string result;

            // make the POST
            using (WebClient client = new WebClient())
            {
                client.UseDefaultCredentials = true;
                client.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";
                byte[] bytes = client.UploadValues(url, "POST", parms);
                result = Encoding.UTF8.GetString(bytes);
            }

            // extract the data
            JObject results = JObject.Parse(result);
            ticket = (string)results["ticket"];
            Globals.RestConnectionOpened = true;
            expires = DateTime.Now.AddMinutes(expiry);
        }

        internal static Int64 CreateFromTemplate(Int64 TemplateID, Int64 ParentID, Int64 ClassificationID, String Name, String Description)
        {
            // update the ticket if needed
            CheckConnection();

            // build the url and parameters
            string url = String.Format("{0}doctemplates/{1}/instances/", Globals.RestUrl, TemplateID);
            NameValueCollection parms = new NameValueCollection();
            parms.Add("parent_id", ParentID.ToString());
            parms.Add("classification_id", ClassificationID.ToString());
            parms.Add("name", Name);
            parms.Add("description", Description);
            string result;

            // make the POST
            using (WebClient client = new WebClient())
            {
                client.UseDefaultCredentials = true;
                client.Headers.Add("otcsticket", ticket);
                client.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";
                byte[] bytes = client.UploadValues(url, "POST", parms);
                result = Encoding.UTF8.GetString(bytes);
            }

            // extract the data
            JObject results = JObject.Parse(result);
            return (Int64)results["node_id"];
        }
    }
}
