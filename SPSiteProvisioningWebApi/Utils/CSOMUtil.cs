
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SPSiteProvisioningWebApi.Utils
{
    public class CSOMUtil
    {
        public static List GetListByTitle(ClientContext clientContext, string listTitle)
        {
            // Get the list collection for the website
            var listColl = clientContext.Web.Lists;

            var q = clientContext.LoadQuery<List>(listColl.Where(n => n.Title == listTitle));
            clientContext.ExecuteQuery();

            if (q.Count() > 0)
            {
                return q.FirstOrDefault<List>();
            }
            else
            {
                return null;
            }
        }
    }
}