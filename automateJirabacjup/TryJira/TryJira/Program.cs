using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.IO;
using System.Web;
using System.Runtime.Serialization.Json;
using System.Web.Script.Serialization;
using System.Configuration;
using Microsoft.Exchange.WebServices.Data;
namespace TryJira
{
    class Program
    {
        static void Main(string[] args)
        {
            string strsummary = null;
            string strdescription = null;
            string strissuetype = null;
            Item strattachmentloc = null;

            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);

            service.Credentials = new NetworkCredential("gades", "suDha@7devi", "camp");

            service.AutodiscoverUrl("sarathchandra.gadepalli@campsystems.com", (a) =>
            {
                return true;
            });

            FindItemsResults<Item> findResults = service.FindItems(
               WellKnownFolderName.Inbox,
               new ItemView(1));

            foreach (Item item in findResults.Items)
            {
                Console.WriteLine(item.Subject);
                strsummary = item.Subject;
                //strdescription = item.Body;
                strissuetype = "Bug";
                strattachmentloc = item;

            }
                

            Jira objJira = new Jira();
            objJira.Url= "http://localhost:8080";
            //string str = File.ReadAllText(@"C:/Users/sarathg/Desktop/sample.txt");
            //string[] words = str.Split(',');

            //foreach (string word in words)
            //{
            //    switch (word.Split(':')[0])
            //    {
            //        case "summary": strsummary = word.Split(':')[1];
            //            break;
            //        case "description": strdescription = word.Split(':')[1];
            //            break;
            //        case "IssueType": strissuetype = word.Split(':')[1];
            //            break;

            //    }

            //}

            JiraJson js = new JiraJson
            {
                fields = new Fields
             {
                 summary = strsummary,
                 //description = strdescription,
                 project = new Project { key = "JIR" },
                 issuetype = new IssueType { name = strissuetype }

             }
            };
            var javaScriptSerializer = new
           System.Web.Script.Serialization.JavaScriptSerializer();
            objJira.JsonString = javaScriptSerializer.Serialize(js);
            objJira.UserName = ConfigurationManager.AppSettings["JiraUserName"];
            objJira.Password = ConfigurationManager.AppSettings["JiraPassword"];
            objJira.filePaths = new List<string>() { "" };
            objJira.AddJiraIssue();
            Console.ReadKey();

        }
    }
    class JiraJson
    {
        public Fields fields { get; set; }
    }
    class Fields
    {
        public string summary { get; set; }
        public string description { get; set; }
        public Project project { get; set; }
        public IssueType issuetype { get; set; }
    }


}
