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

        public static void CreateTimerJob(SPSite site)
        {
            var timerJob = new PowiadomieniaTimerJob(site);
            timerJob.Schedule = new SPMinuteSchedule
            {
                BeginSecond = 0,
                EndSecond = 0
            };
            timerJob.Update();
        }

        public static void DelteTimerJob(SPSite site)
        {
            site.WebApplication.JobDefinitions
                .OfType<PowiadomieniaTimerJob>()
                .Where(i => string.Equals(i.SiteUrl, site.Url, StringComparison.InvariantCultureIgnoreCase))
                .ToList()
                .ForEach(i => i.Delete());
        }

        public PowiadomieniaTimerJob()
            : base()
        {

        }

        public PowiadomieniaTimerJob(SPSite site)
            : base(string.Format("Bizarre_Powiadomienia Update Timer Job ({0})", site.Url), site.WebApplication, null, SPJobLockType.Job)
        {
            Title = Name;
            SiteUrl = site.Url;
        }

        public string SiteUrl
        {
            get { return (string)this.Properties["SiteUrl"]; }
            set { this.Properties["SiteUrl"] = value; }
        }

        public override void Execute(Guid targetInstanceId)
        {

            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                using (var site = new SPSite(SiteUrl))
                {
                    var targetList = site.RootWeb.Lists.TryGetList("Powiadomienia");

                    if (targetList!=null)
                    {
                        targetList.Items.Cast<SPListItem>()
                            .Where(i => (bool)i["Wys_x0142_ane"] != true)
                            .ToList()
                            .ForEach(item =>
                                {
                                    try
                                    {
                                        item["Title"] = string.Concat("TimerJob : ", DateTime.Now.ToString());
                                        item.Update();

                                        StartWorkflow(item, "Powiadomienie.OnUpdate");

                                    }
                                    catch (Exception)
                                    { }
                                });
                    }

                }
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




