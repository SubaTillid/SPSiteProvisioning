using Microsoft.SharePoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
namespace SPSiteProvisioning
{
    class Program
    {
        static void Main(string[] args)
        {
            var userName = ConfigurationSettings.AppSettings["username"];
            var password = ConfigurationSettings.AppSettings["password"];
            Console.WriteLine("userName" + userName);
            Console.WriteLine("password" + password);
            Console.ReadLine();            
        }

        public void createSite()
        {

        }
        
    }

    
}
