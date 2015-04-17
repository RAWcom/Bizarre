using System;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Security;
using System.Linq;

namespace Bizarre.PowiadomieniaTimerJob.Features.PowiadomieniaTimerJobFeature
{
    /// <summary>
    /// This class handles events raised during feature activation, deactivation, installation, uninstallation, and upgrade.
    /// </summary>
    /// <remarks>
    /// The GUID attached to this class may be used during packaging and should not be modified.
    /// </remarks>

    [Guid("9e5c11f8-11ee-423a-968b-d1b2a7831dbb")]
    public class PowiadomieniaTimerJobFeatureEventReceiver : SPFeatureReceiver
    {

        private const string List_JOB_NAME = "PowiadomieniaTimerJob"; 

        // Uncomment the method below to handle the event raised after a feature has been activated.

        public override void FeatureActivated(SPFeatureReceiverProperties properties)
        {
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                SPSite site = properties.Feature.Parent as SPSite;

                // make sure the job isn't already registered 
                site.WebApplication.JobDefinitions.Where(t => t.Name.Equals(List_JOB_NAME)).ToList().ForEach(j => j.Delete());

                // install the job 
                PowiadomieniaTimerJob listLoggerJob = new PowiadomieniaTimerJob(List_JOB_NAME, site.WebApplication);
                SPMinuteSchedule schedule = new SPMinuteSchedule();
                schedule.BeginSecond = 0;
                schedule.EndSecond = 59;
                schedule.Interval = 5;
                listLoggerJob.Schedule = schedule;
                listLoggerJob.Update();
            }); 
        }


        // Uncomment the method below to handle the event raised before a feature is deactivated.

        public override void FeatureDeactivating(SPFeatureReceiverProperties properties)
        {
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                SPSite site = properties.Feature.Parent as SPSite;

                // delete the job 
                site.WebApplication.JobDefinitions.Where(t => t.Name.Equals(List_JOB_NAME)).ToList().ForEach(j => j.Delete());
            }); 
        }


        // Uncomment the method below to handle the event raised after a feature has been installed.

        //public override void FeatureInstalled(SPFeatureReceiverProperties properties)
        //{
        //}


        // Uncomment the method below to handle the event raised before a feature is uninstalled.

        //public override void FeatureUninstalling(SPFeatureReceiverProperties properties)
        //{
        //}

        // Uncomment the method below to handle the event raised when a feature is upgrading.

        //public override void FeatureUpgrading(SPFeatureReceiverProperties properties, string upgradeActionName, System.Collections.Generic.IDictionary<string, string> parameters)
        //{
        //}
    }
}
