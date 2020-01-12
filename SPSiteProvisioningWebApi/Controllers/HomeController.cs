using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Application;
using SharePointPnP.IdentityModel;
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
using OfficeDevPnP.Core.ALM;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Pages;
using OfficeDevPnP.Core.Diagnostics;
using Newtonsoft.Json.Linq;

namespace SPSiteProvisioningWebApi.Controllers
{
    public class HomeController : Controller
    {
        private static ClientContext templateSiteClientContext;

        private static ClientContext binderSiteClientContext;

        private static string templateSiteUrl;

        private static string binderSiteUrl;

        private static SecureString passWord;

        private static string userName;

        public ActionResult Index()
        {
            ViewBag.Title = "Home Page";
            userName = "murali@chennaitillidsoft.onmicrosoft.com";
            passWord = new SecureString();
            foreach (char c in "n@#eTD@098!".ToCharArray())
            {
                passWord.AppendChar(c);
            };
            templateSiteUrl = "https://chennaitillidsoft.sharepoint.com/sites/developer5";
            binderSiteUrl = "https://chennaitillidsoft.sharepoint.com/sites/POC/SiteProvisioning";
            AddSectionAndAddWebpart();
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
            WorkflowDefinition currentWorkFlow = publishedWorkflowDefinitions.Where(flow => flow.DisplayName.Equals("BBH Document Atestation")).First();
            ProvisionWorkFlowAndRelatedList(currentWorkFlow.Xaml, binderSiteUrl);
        }

        public static void ProvisionWorkFlowAndRelatedList(string workFlowXMlFile, string binderSiteUrl)
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
            var bbhDocumentWFDefinitionId = service.SaveDefinitionAndPublish("BBHDocument", bbhDocumentWF);

            //Create the workflow tasks list
            //var taskListId = service.CreateTaskList("BBHDocument Workflow Tasks");
            var taskListId = CSOMUtil.GetListByTitle(binderSiteClientContext, "BBHDocument Workflow Tasks");
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
            using (binderSiteClientContext = new ClientContext(binderSiteUrl))
            {
                binderSiteClientContext.Credentials = new SharePointOnlineCredentials("murali@chennaitillidsoft.onmicrosoft.com", passWord);
                
                //Guid spfxExtension_GlobalHeaderID = new Guid("1e3d3ef7-0983-4d40-9dbb-9c6d4539639a");
                //string spfxExtName = "react-logo-festoon-client-side-solution";
                //string spfxExtTitle = "LogoFestoonApplicationCustomizer";

                //string spfxExtDescription = "Logo Festoon Application Customizer";
                //string spfxExtLocation = "ClientSideExtension.ApplicationCustomizer";
                ////string spfxExtProps = "";  // add properties if any, else remove this

                //UserCustomAction userCustomAction = binderSiteClientContext.Site.UserCustomActions.Add();
                //userCustomAction.Name = spfxExtName;
                //userCustomAction.Title = spfxExtTitle;
                //userCustomAction.Description = spfxExtDescription;
                //userCustomAction.Location = spfxExtLocation;
                //userCustomAction.ClientSideComponentId = spfxExtension_GlobalHeaderID;
                ////userCustomAction.ClientSideComponentProperties = spfxExtProps;

                //binderSiteClientContext.ExecuteQuery();
                //using (binderSiteClientContext = new ClientContext(binderSiteUrl))
                //{

                    var appManager = new AppManager(binderSiteClientContext); 
                    var apps = appManager.GetAvailable(); 
                    var chartsApp = apps.Where(a => a.Title == "react-logo-festoon-client-side-solution").FirstOrDefault(); 
                    var installApp = appManager.Install(chartsApp);
                    if (installApp)
                    {
                        Guid spfxExtension_GlobalHeaderID1 = chartsApp.Id; 
                        string spfxExtName1 = chartsApp.Title; 
                        string spfxExtTitle1 = chartsApp.Title; 
                        //string spfxExtGroup1 = ""; 
                        string spfxExtDescription1= "Logo Festoon Application Customizer"; 
                        string spfxExtLocation1 = "ClientSideExtension.ApplicationCustomizer";
                        CustomActionEntity ca = new CustomActionEntity
                        {
                            Name = spfxExtName1,
                            Title = spfxExtTitle1,
                            //Group = spfxExtGroup1,
                            Description = spfxExtDescription1,
                            Location = spfxExtLocation1,
                            ClientSideComponentId = spfxExtension_GlobalHeaderID1
                        };

                        binderSiteClientContext.Web.AddCustomAction(ca);
                        binderSiteClientContext.ExecuteQueryRetry();
                    }
                }
            }
        
        public static void PageSectionDivision()
        {
            OfficeDevPnP.Core.AuthenticationManager authenticationManager = new OfficeDevPnP.Core.AuthenticationManager();
            using(ClientContext currentSiteContext = authenticationManager.GetSharePointOnlineAuthenticatedContextTenant(binderSiteUrl, userName, passWord)) {
                string pageName = "POCSiteProvisioning.aspx";
                ClientSidePage page = ClientSidePage.Load(currentSiteContext, pageName);
                var appManager = new AppManager(currentSiteContext);
                var apps = appManager.GetAvailable();
                ClientSideComponent clientSideComponent = null;
                var chartsApp = apps.Where(a => a.Title == "hero-control-client-side-solution-ProductionEnv").FirstOrDefault();
                bool controlPresent = false;
                bool done = false;
                int count = 0;
                do
                {
                    try
                    {
                        ClientSideComponent[] clientSideComponents = (ClientSideComponent[])page.AvailableClientSideComponents();
                        clientSideComponent = clientSideComponents.Where(c => c.Id.ToLower() == chartsApp.Id.ToString().ToLower()).FirstOrDefault();
                        foreach (ClientSideComponent csc in clientSideComponents)
                        {
                            if (csc.Id.ToString().ToLower().Contains(chartsApp.Id.ToString().ToLower()))
                            {
                                clientSideComponent = csc;
                                continue;
                            }
                        }
                        //ClientSideWebPart webPart = page.Controls.Where(wP => wP != null)
                        foreach (var control in page.Controls)
                        {
                            ClientSideWebPart cpWP = control as ClientSideWebPart;
                            if (cpWP != null && cpWP.SpControlData.WebPartId.ToString() == chartsApp.Id.ToString())
                            {
                                controlPresent = true;
                                done = true;
                            }
                        }

                        if (!controlPresent)
                        {

                            ClientSideWebPart WebPart = new ClientSideWebPart(clientSideComponent);
                            JToken activeValueToken = true;

                            // Find the web part configuration string from the web part file or code debugging
                            //string propertyJSONString = String.Format("[{{<WP Configuration string>}}]", < parameters >);
                            //JToken propertyTermToken = JToken.Parse(propertyJSONString);
                            WebPart.Properties.Add("showOnlyActive", activeValueToken);

                            CanvasSection section = new CanvasSection(page, CanvasSectionTemplate.ThreeColumnVerticalSection, page.Sections.Count + 1);
                            page.Sections.Add(section);
                            page.Save();
                            page.AddControl(WebPart, section.Columns[0]);
                            page.Save();
                            page.Publish();
                            done = true;
                            controlPresent = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        //Log.Info("Catched exception while adding Capex web part.. Trying again" + ex.Message);
                        Console.WriteLine(ex);
                        Console.ReadLine();
                        count++;
                    }
                } while (!done && count <= 5);
            }
        }
        
        public static void AddSectionAndAddWebpart()
        {
            OfficeDevPnP.Core.AuthenticationManager authenticationManager = new OfficeDevPnP.Core.AuthenticationManager();
            using (binderSiteClientContext = authenticationManager.GetSharePointOnlineAuthenticatedContextTenant(binderSiteUrl, userName, passWord))
            {
                //Create a page or get the existing page
                string pageName = "POCSiteProvisioning.aspx";
                ClientSidePage page = ClientSidePage.Load(binderSiteClientContext, pageName);
                //var page = binderSiteClientContext.Web.AddClientSidePage("POCAppProvisioning.aspx", true);

                // Add Section 
                page.AddSection(CanvasSectionTemplate.ThreeColumn, 5);

                // get the available web parts - this collection will include OOTB and custom SPFx web parts..
                page.Save();

                // Get all the available webparts
                var components = page.AvailableClientSideComponents();

                // add the named web part..
                var webPartToAdd = components.Where(wp => wp.ComponentType == 1 && wp.Name == "HeroControl").FirstOrDefault();

                if (webPartToAdd != null)
                {
                    ClientSideWebPart clientWp = new ClientSideWebPart(webPartToAdd) { Order = 1 };

                    //Add the WebPart to the page with appropriate section
                    page.AddControl(clientWp, page.Sections[1].Columns[1]);

                }

                // the save method creates the page if one doesn't exist with that name in this site..
                page.Save();
            }
        }
    }

}


