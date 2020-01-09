using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WorkflowServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Web;
using System.Web.Mvc;
using Microsoft.SharePoint.Client.Publishing;
using SPSiteProvisioningWebApi.Models;
using System.IO;

namespace SPSiteProvisioningWebApi.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            ViewBag.Title = "Home Page";
            getAllWorkFlows();
            return View();
        }
        public ActionResult TestApp()
        {
            ViewBag.Title = "Home Page";
            getAllWorkFlows();
            return View();
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
            WorkflowDefinition currentWorkFlow = publishedWorkflowDefinitions.Where(flow => flow.DisplayName.Equals("BBH Document Atestation")).First();
            if(currentWorkFlow != null)
            {
                var workFlowTemplate = currentWorkFlow;
                var workflowSubscriptionService = workflowServicesManager.GetWorkflowSubscriptionService();

                // get all workflow associations
                var workflowAssociations = workflowSubscriptionService.EnumerateSubscriptionsByDefinition(currentWorkFlow.Id);
                context.Load(workflowAssociations);
                context.ExecuteQuery();

                foreach (var association in workflowAssociations)
                {
                    Console.WriteLine("{0} - {1}",
                      association.Id, association.Name);

                }
                var binderSiteUrl = "https://chennaitillidsoft.sharepoint.com/sites/POC/spotlight/HeroControlDev/";
                ClientContext clientContext = new ClientContext(binderSiteUrl);
                clientContext.Credentials = new SharePointOnlineCredentials("murali@chennaitillidsoft.onmicrosoft.com", passWord);


                //Construct object with workflow template info
                WorkflowTemplateInfo solutionInfo = new WorkflowTemplateInfo();
                var solutionPath = "../../SiteAssets/BBH Document Atestation.wsp";
                solutionInfo.PackageFilePath = solutionPath;
                //PackageName is mandatory
                solutionInfo.PackageName = Path.GetFileNameWithoutExtension(solutionPath);
                //Guid is automatically predefined in template file (.wsp)
                solutionInfo.PackageGuid = workFlowTemplate.Id;
                //Workflow feature Id is need to activate workflow in the web
                //solutionInfo.FeatureId = workFlowTemplate.;
                //Init workflow template deployer
                using (AddWorkFlowFormExistingTemplate workflowDeployer = new AddWorkFlowFormExistingTemplate(context))
                {
                    //Provisioning workflow resources
                    workflowDeployer.DeployWorkflowSolution(solutionPath);
                    //Activates workflow template
                    workflowDeployer.ActivateWorkflowSolution(solutionInfo);
                }
            }
            Console.ReadLine();
        }
    }
}
