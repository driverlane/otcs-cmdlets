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
            // details
            string url = String.Format("{0}auth/", Globals.RestUrl);
            NameValueCollection parms = new NameValueCollection();
            parms.Add("username", "a");
            parms.Add("password", "a");
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


    }
}
