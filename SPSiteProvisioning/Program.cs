using Microsoft.SharePoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using Microsoft.SharePoint.Client;
using Microsoft.Online.SharePoint.Client.TenantAdmin;
using Microsoft.Online.SharePoint.TenantAdministration;
using System.Security;
using OfficeDevPnP.Core.Sites;
using OfficeDevPnP.Core.Utilities;
using OfficeDevPnP.Core;
using Microsoft.SharePoint.Client.WorkflowServices;


//using Microsoft.
namespace SPSiteProvisioning
{
    public class Program
    {
        private static ClientContext clientContext;
        private static string siteUrl;
        private static string appID;
        private static string appSecret;


        public static void Main(string[] args)
        {
            getAllWorkFlows();
        }

        public static async void createSite()
        {
            try
            {
                AuthenticationManager authManager = new AuthenticationManager();
                clientContext = authManager.GetAppOnlyAuthenticatedContext(siteUrl, appID, appSecret);
                var teamContext = await clientContext.CreateSiteAsync(
                        new TeamSiteCollectionCreationInformation
                        {
                            Alias = "TestSiteProvision", // Mandatory
                            DisplayName = "TestSiteProvision", // Mandatory
                            Description = "Testing SiteProvision CSOM", // Optional
                            Classification = "classification", // Optional
                            IsPublic = true, // Optional, default true
                        });
                teamContext.Load(teamContext.Web, w => w.Url);
                teamContext.ExecuteQueryRetry();
                Console.WriteLine(teamContext.Web.Url);
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }


        public static void getAllWorkFlows()
        {
            var siteURl = "https://chennaitillidsoft.sharepoint.com/sites/developer5";
            var context = new ClientContext(siteURl);
            SecureString passWord = new SecureString();
            foreach (char c in "n@#eTD@098!".ToCharArray()) passWord.AppendChar(c);
            context.Credentials = new SharePointOnlineCredentials("murali@chennaitillidsoft.onmicrosoft.com", passWord);

            var workflowServicesManager = new WorkflowServicesManager(context, context.Web);

            // connect to the deployment service 
            var workflowDeploymentService = workflowServicesManager.GetWorkflowDeploymentService();

            // get all installed workflows
            var publishedWorkflowDefinitions = workflowDeploymentService.EnumerateDefinitions(true);
            context.Load(publishedWorkflowDefinitions);
            context.ExecuteQuery();

            // display list of all installed workflows
            foreach (var workflowDefinition in publishedWorkflowDefinitions)
            {
                Console.WriteLine("{0} - {1}", workflowDefinition.Id.ToString(), workflowDefinition.DisplayName);
            }
            Console.ReadLine();            
        }
    }
}

