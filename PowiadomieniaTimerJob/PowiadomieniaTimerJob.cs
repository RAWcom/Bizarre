using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Workflow;

namespace Bizarre.PowiadomieniaTimerJob
{

    public class PowiadomieniaTimerJob : Microsoft.SharePoint.Administration.SPJobDefinition
    {
        /// <summary>
        /// wskazuje który site ma zostać wybrany jako domyślny, koniecznie należy go przestawić przed wgraniem na produkcję
        /// </summary>
        /// <remarks>
        /// true = sites["sites/bw"]
        /// false = sites[0]
        /// </remarks>
        bool DEV_MODE = false; 

       
        /// <summary> 
        /// Default Consructor 
        /// </summary> 
        public PowiadomieniaTimerJob()
            : base()
        {
        }

        /// <summary> 
        /// Parameterized Constructor 
        /// </summary> 
        /// <param name="jobName">Name of Job to display in central admin</param> 
        /// <param name="service">SharePoint Service </param> 
        /// <param name="server">Name of the server</param> 
        /// <param name="targetType">job type is for content db or job</param> 
        public PowiadomieniaTimerJob(string jobName, SPService service, SPServer server, SPJobLockType targetType)
            : base(jobName, service, server, targetType)
        {
        }

        /// <summary> 
        /// Parameterized Constructor 
        /// </summary> 
        /// <param name="jobName"></param> 
        /// <param name="webApplication"></param> 
        public PowiadomieniaTimerJob(string jobName, SPWebApplication webApplication)
            : base(jobName, webApplication, null, SPJobLockType.ContentDatabase)
        {
            this.Title = "Bizarre_Powiadomienia Timer Job";
        }

        public override void  Execute(Guid contentDbId)
        {
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                
                // get a reference to the current site collection's content database 
                SPWebApplication webApplication = this.Parent as SPWebApplication;
                SPContentDatabase contentDb = webApplication.ContentDatabases[contentDbId];

                SPWeb rootWeb;
                
                if (DEV_MODE)
                {
                    rootWeb = contentDb.Sites["sites/BW"].RootWeb; //wartość developerska
                }
                else
                {
                    rootWeb = contentDb.Sites[0].RootWeb;
                }
                
                SPList list = rootWeb.Lists.TryGetList("Powiadomienia");


                //if (list!=null)
                //{
                //    //SPListItem li = list.AddItem();
                //    //li["Title"] = DateTime.Now.ToString();
                //    //li.Update();

                //    SPListItem newListItem = list.Items.Add();
                //    newListItem["Title"] = string.Concat("Powiadomienie : ", DateTime.Now.ToString());
                //    newListItem.Update();
                //}


                if (list != null)
                {
                    //dla wszystkich rekordów o statusie niewysłane
                    StringBuilder sb = new StringBuilder(@"<OrderBy><FieldRef Name='ID' /></OrderBy><Where><Neq><FieldRef Name='Wys_x0142_ane' /><Value Type='Boolean'>1</Value></Neq></Where>");

                    string camlQuery = sb.ToString();

                    SPQuery query = new SPQuery();
                    query.Query = camlQuery;

                    SPListItemCollection items = list.GetItems(query);
                    //SPListItemCollection items = list.GetItems();

                    if (items.Count > 0)
                    {
                        foreach (SPListItem item in items)
                        {
                            try
                            {
                                item["Title"] = string.Concat("TimerJob : ", DateTime.Now.ToString());
                                item.Update();

                                StartWorkflow(item, "Powiadomienie.OnUpdate");
    
                            }
                            catch (Exception)
                            { }

                        }
                    }
                }

                //rootWeb.Dispose();
            });
        }

        #region Helpers

        private static void StartWorkflow(SPListItem listItem, string workflowName)
        {
            try
            {
                SPWorkflowManager manager = listItem.Web.Site.WorkflowManager;
                SPWorkflowAssociationCollection objWorkflowAssociationCollection = listItem.ParentList.WorkflowAssociations;
                foreach (SPWorkflowAssociation objWorkflowAssociation in objWorkflowAssociationCollection)
                {
                    if (String.Compare(objWorkflowAssociation.Name, workflowName, true) == 0)
                    {

                        //We found our workflow association that we want to trigger.

                        //Replace the workflow_GUID with the GUID of the workflow feature that you
                        //have deployed.

                        try
                        {
                            manager.StartWorkflow(listItem, objWorkflowAssociation, objWorkflowAssociation.AssociationData, true);
                            //The above line will start the workflow...
                        }
                        catch (Exception)
                        { }


                        break;
                    }
                }
            }
            catch (Exception)
            {}
        }

        #endregion
    }
}




