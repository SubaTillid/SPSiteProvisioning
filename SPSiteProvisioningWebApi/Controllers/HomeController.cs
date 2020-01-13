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
using SPSiteProvisioningWebApi.Utils;
using SPSiteProvisioningWebApi.Services;

namespace SPSiteProvisioningWebApi.Controllers
{
    public class HomeController : Controller
    {
        private static ClientContext templateSiteClientContext;

        private static ClientContext binderSiteClientContext;

        private static string templateSiteUrl;

        private static string binderSiteUrl;

        private static SecureString passWord;

        public ActionResult Index()
        {
            ViewBag.Title = "Home Page";
            passWord = new SecureString();
            foreach (char c in "n@#eTD@098!".ToCharArray())
            {
                passWord.AppendChar(c);
            };
            templateSiteUrl = "https://chennaitillidsoft.sharepoint.com/sites/developer5";
            binderSiteUrl = "https://chennaitillidsoft.sharepoint.com/sites/POC/SiteProvisioning";
            SetPropertyBagValue();
            return View();
        }
        public ActionResult TestApp()
        {
            ViewBag.Title = "Home Page";
            passWord = new SecureString();
            foreach (char c in "n@#eTD@098!".ToCharArray())
            {
                passWord.AppendChar(c);
            };
            ProvisionWorkFlow();
            return View();
        }
        public static void ProvisionWorkFlow()
        {
            templateSiteClientContext = new ClientContext(templateSiteUrl);
            
            templateSiteClientContext.Credentials = new SharePointOnlineCredentials("murali@chennaitillidsoft.onmicrosoft.com", passWord);

            var workflowServicesManager = new WorkflowServicesManager(templateSiteClientContext, templateSiteClientContext.Web);

            // connect to the deployment service 
            var workflowDeploymentService = workflowServicesManager.GetWorkflowDeploymentService();

            // get all installed workflows
            var publishedWorkflowDefinitions = workflowDeploymentService.EnumerateDefinitions(true);
            templateSiteClientContext.Load(publishedWorkflowDefinitions);
            templateSiteClientContext.ExecuteQuery();

            // display list of all installed workflows
            WorkflowDefinition currentWorkFlow = publishedWorkflowDefinitions.Where(flow => flow.DisplayName.Equals("BBH Document Atestation")).First();
            if (currentWorkFlow != null)
            {
                var workFlowTemplate = currentWorkFlow;
                //var workflowSubscriptionService = workflowServicesManager.GetWorkflowSubscriptionService();

                //// get all workflow associations
                //var workflowAssociations = workflowSubscriptionService.EnumerateSubscriptionsByDefinition(currentWorkFlow.Id);
                //templateSiteClientContext.Load(workflowAssociations);
                //templateSiteClientContext.ExecuteQuery();

                //foreach (var association in workflowAssociations)
                //{
                //    Console.WriteLine("{0} - {1}",
                //      association.Id, association.Name);

                //}
                //var binderSiteUrl = "https://chennaitillidsoft.sharepoint.com/sites/POC/spotlight/HeroControlDev/";
                binderSiteClientContext = new ClientContext(binderSiteUrl);
                binderSiteClientContext.Credentials = new SharePointOnlineCredentials("murali@chennaitillidsoft.onmicrosoft.com", passWord);


                //Construct object with workflow template info
                WorkflowTemplateInfo solutionInfo = new WorkflowTemplateInfo();
                var solutionPath = "https://chennaitillidsoft.sharepoint.com/sites/developer5/SiteAssets/BBH Document Atestation.wsp";
                solutionInfo.PackageFilePath = solutionPath;
                //PackageName is mandatory
                solutionInfo.PackageName = Path.GetFileNameWithoutExtension(solutionPath);
                //Guid is automatically predefined in template file (.wsp)
                solutionInfo.PackageGuid = workFlowTemplate.Id;
                //Workflow feature Id is need to activate workflow in the web
                //solutionInfo.FeatureId = workFlowTemplate.;
                //Init workflow template deployer
                using (AddWorkFlowFormExistingTemplate workflowDeployer = new AddWorkFlowFormExistingTemplate(binderSiteClientContext))
                {
                    //Provisioning workflow resources
                    workflowDeployer.DeployWorkflowSolution(solutionPath);
                    //Activates workflow template
                    workflowDeployer.ActivateWorkflowSolution(solutionInfo);
                }
            }
            Console.ReadLine();
        }

        public static void GetWorkFlowTemplate()
        {
            var templateSiteClientContext = new ClientContext(templateSiteUrl);

            templateSiteClientContext.Credentials = new SharePointOnlineCredentials("murali@chennaitillidsoft.onmicrosoft.com", passWord);

            var workflowServicesManager = new WorkflowServicesManager(templateSiteClientContext, templateSiteClientContext.Web);

            // connect to the deployment service 
            var workflowDeploymentService = workflowServicesManager.GetWorkflowDeploymentService();

            // get all installed workflows
            var publishedWorkflowDefinitions = workflowDeploymentService.EnumerateDefinitions(true);
            templateSiteClientContext.Load(publishedWorkflowDefinitions);
            templateSiteClientContext.ExecuteQuery();

            // display list of all installed workflows
            WorkflowDefinition currentWorkFlow = publishedWorkflowDefinitions.Where(flow => flow.DisplayName.Equals("BBH Document Atestation")).FirstOrDefault();
            ProvisionWorkFlowAndRelatedList(currentWorkFlow.Xaml, binderSiteUrl);
        }

        public static void ProvisionWorkFlowAndRelatedList(string workFlowXMlFile , string binderSiteUrl)
        {
            //Return the list the workflow will be associated with
            binderSiteClientContext = new ClientContext(binderSiteUrl);
            binderSiteClientContext.Credentials = new SharePointOnlineCredentials("murali@chennaitillidsoft.onmicrosoft.com", passWord);
            var bbhDocumentList = CSOMUtil.GetListByTitle(binderSiteClientContext, "BBH Documents");

            //Create a new WorkflowProvisionService class instance which uses the 
            //Microsoft.SharePoint.Client.WorkflowServices.WorkflowServicesManager to
            //provision and configure workflows
            var service = new WorkflowProvisionService(binderSiteClientContext);

            //Read the workflow .XAML file
            //var bbhDocumentWF = System.IO.File.ReadAllText(workFlowXMlFile);
            //string invalid = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
            //foreach (char c in invalid)
            //{
            //    workFlowXMlFile = workFlowXMlFile.Replace(c.ToString(), "");
            //}
            //var solutionPath = "E:/Site Provisioning files/BBH Document Atestation (1).wsp";
            var bbhDocumentWF = workFlowXMlFile;


            //Create the WorkflowDefinition and use the 
            //Microsoft.SharePoint.Client.WorkflowServices.WorkflowServicesManager
            //to save and publish it.  
            //This method is shown below for reference.
            //var bbhDocumentWFDefinitionId = service.SaveDefinitionAndPublish("BBHDocument", WorkflowUtil.TranslateWorkflow(bbhDocumentWF));
            var bbhDocumentWFDefinitionId = service.SaveDefinitionAndPublish("BBHDocument",bbhDocumentWF);

            //Create the workflow tasks list
            //var taskListId = service.CreateTaskList("BBHDocument Workflow Tasks");
            var taskListId =  CSOMUtil.GetListByTitle(binderSiteClientContext, "BBHDocument Workflow Tasks");
            //Create the workflow history list
            var historyListId = service.CreateHistoryList("BBHDocument Workflow History");

            //Use the Microsoft.SharePoint.Client.WorkflowServices.WorkflowSubscriptionService to 
            //subscibe the workflow to the list the workflow is associated with, register the
            //events it is associated with, and register the tasks and history list. 
            //This method is shown below for reference.
            service.Subscribe("BBHDocument Workflow", bbhDocumentWFDefinitionId, bbhDocumentList.Id,
                WorkflowSubscritpionEventType.ItemAdded, taskListId.Id, historyListId);
        }

        public static void SetPropertyBagValue()
        {
            using (binderSiteClientContext = new ClientContext(binderSiteUrl))
            {
                binderSiteClientContext.Credentials = new SharePointOnlineCredentials("murali@chennaitillidsoft.onmicrosoft.com", passWord);
                binderSiteClientContext.Web.SetPropertyBagValue("Test Update Property Bag Value 1", "Successfully Updated");
                binderSiteClientContext.ExecuteQuery();
                var propertyBagValue = binderSiteClientContext.Web.GetPropertyBagValueString("Test Update Property Bag Value", "Not Found");
            }
        }

        public static void AddSPFXExtension()
        {
            using(binderSiteClientContext = new ClientContext(binderSiteUrl))
            {
                binderSiteClientContext.Credentials = new SharePointOnlineCredentials("murali@chennaitillidsoft.onmicrosoft.com", passWord);

                //AppDeclaration App = appConfiguration.Apps.Find(app => app.appName.Contains("<app name>"));
                Guid spfxExtension_GlobalHeaderID = new Guid("59a815be-4478-4ca7-b992-3c42fd0bdfaf");
                string spfxExtName = "react-logo-festoon";
                string spfxExtTitle = "Application Extension - Deployment of custom action.";

                string spfxExtDescription = "Deploys a custom action with ClientSideComponentId association";
                string spfxExtLocation = "ClientSideExtension.ApplicationCustomizer";
                string spfxExtProps = "";  // add properties if any, else remove this

                UserCustomAction userCustomAction = binderSiteClientContext.Site.UserCustomActions.Add();
                userCustomAction.Name = spfxExtName;
                userCustomAction.Title = spfxExtTitle;
                userCustomAction.Description = spfxExtDescription;
                userCustomAction.Location = spfxExtLocation;
                userCustomAction.ClientSideComponentId = spfxExtension_GlobalHeaderID;
                userCustomAction.ClientSideComponentProperties = spfxExtProps;

                binderSiteClientContext.Site.Context.ExecuteQuery();
            }
        }
    }
}

