using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SPSiteProvisioningWebApi.Models
{
    public class WorkflowTemplateInfo
    {
        /// <summary>
        /// Path to the workflow template(wsp)
        /// </summary>
        public string PackageFilePath { get; set; }
        /// <summary>
        /// Globally unique identifier for workflow template
        /// </summary>
        public Guid PackageGuid { get; set; }
        /// <summary>
        /// Workflow user solution name
        /// </summary>
        public string PackageName { get; set; }
        /// <summary>
        /// Web-scoped worfklow feature Id for activating worfklow
        /// </summary>
        public Guid FeatureId { get; set; }
    }
}